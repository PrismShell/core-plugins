using System.Runtime.InteropServices;
using System.Text.Json;
using Plugboard.Contracts;

namespace Plugboard.Plugins.Outlook;

// Ported from the local-gateway OutlookHandler (COM automation via dynamic).
// Each route returns its data object (host wraps in { ok, data }); errors throw
// (host wraps in { ok:false, error }). GetStatus is the exception: "Outlook not
// running" is reported as data, not an error.
public sealed class OutlookPlugin : IPlugin
{
    public string Name => "outlook";

    public void Register(IEndpointRegistry r)
    {
        r.Map("GET",  "outlook/status",           _ => Task.FromResult<object?>(GetStatus()));
        r.Map("GET",  "outlook/folders",          _ => Task.FromResult<object?>(GetFolders()));
        r.Map("POST", "outlook/send",             q => Json(q, SendEmail));
        r.Map("POST", "outlook/inbox",            q => Json(q, GetInbox));
        r.Map("POST", "outlook/email",            q => Json(q, GetEmail));
        r.Map("POST", "outlook/search",           q => Json(q, SearchEmails));
        r.Map("POST", "outlook/move",             q => Json(q, MoveEmail));
        r.Map("POST", "outlook/delete",           q => Json(q, DeleteEmail));
        r.Map("POST", "outlook/mark-read",        q => Json(q, MarkRead));
        r.Map("POST", "outlook/reply",            q => Json(q, Reply));
        r.Map("POST", "outlook/forward",          q => Json(q, Forward));
        r.Map("GET",  "outlook/calendar",         q => Task.FromResult<object?>(GetCalendar(q.Query.GetValueOrDefault("startDate"), q.Query.GetValueOrDefault("endDate"))));
        r.Map("POST", "outlook/calendar/create",  q => Json(q, CreateAppointment));
        r.Map("POST", "outlook/calendar/delete",  q => Json(q, DeleteAppointment));
        r.Map("GET",  "outlook/contacts",         q => Task.FromResult<object?>(GetContacts(q.Query.GetValueOrDefault("folder"))));
        r.Map("POST", "outlook/contacts/create",  q => Json(q, CreateContact));
        r.Map("POST", "outlook/contacts/search",  q => Json(q, SearchContacts));
    }

    private static Task<object?> Json(PluginRequest req, Func<JsonElement, object?> handler)
    {
        using var doc = JsonDocument.Parse(string.IsNullOrWhiteSpace(req.Body) ? "{}" : req.Body);
        return Task.FromResult(handler(doc.RootElement));
    }

    private static dynamic GetOutlook() => Activator.CreateInstance(Type.GetTypeFromProgID("Outlook.Application")!)!;
    private static dynamic GetNamespace(dynamic outlook) => outlook.GetNamespace("MAPI");

    private static dynamic GetFolder(dynamic ns, string? folderPath)
    {
        if (string.IsNullOrEmpty(folderPath)) return ns.GetDefaultFolder(6); // olFolderInbox
        var parts = folderPath.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
        dynamic folder = ns.Folders[parts[0]];
        for (int i = 1; i < parts.Length; i++) folder = folder.Folders[parts[i]];
        return folder;
    }

    private static void Release(params dynamic[] objs)
    {
        foreach (var o in objs) try { Marshal.ReleaseComObject(o); } catch { }
    }

    private static object GetStatus()
    {
        try
        {
            dynamic outlook = GetOutlook();
            dynamic ns = GetNamespace(outlook);
            var name = (string)ns.CurrentUser.Name;
            Release(ns, outlook);
            return new { outlookRunning = true, user = name };
        }
        catch (Exception ex) { return new { outlookRunning = false, error = ex.Message }; }
    }

    private static object? GetFolders()
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        var folders = new List<object>();
        void Walk(dynamic folder, int depth)
        {
            folders.Add(new { name = (string)folder.Name, folderPath = (string)folder.FolderPath, itemCount = (int)folder.Items.Count, unreadCount = (int)folder.UnReadItemCount });
            if (depth < 1) foreach (dynamic sub in folder.Folders) Walk(sub, depth + 1);
        }
        foreach (dynamic store in ns.Folders) Walk(store, 0);
        Release(ns, outlook);
        return new { folders };
    }

    private static object? SendEmail(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic mail = outlook.CreateItem(0); // olMailItem
        mail.To = body.GetProperty("to").GetString()!;
        mail.Subject = body.GetProperty("subject").GetString()!;
        mail.Body = body.GetProperty("body").GetString()!;
        if (body.TryGetProperty("cc", out var cc)) mail.CC = cc.GetString()!;
        if (body.TryGetProperty("bcc", out var bcc)) mail.BCC = bcc.GetString()!;
        if (body.TryGetProperty("htmlBody", out var hb)) mail.HTMLBody = hb.GetString()!;
        if (body.TryGetProperty("attachments", out var atts))
            foreach (var att in atts.EnumerateArray()) mail.Attachments.Add(att.GetString()!);

        bool send = body.TryGetProperty("send", out var s) && s.GetBoolean();
        string message;
        if (send) { mail.Send(); message = "Email sent"; }
        else { mail.Display(false); message = "Draft opened in Outlook"; }
        Release(mail, outlook);
        return new { message };
    }

    private static object? GetInbox(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        var folderPath = body.TryGetProperty("folder", out var fp) ? fp.GetString() : null;
        dynamic folder = GetFolder(ns, folderPath);
        int count = body.TryGetProperty("count", out var c) ? c.GetInt32() : 25;

        dynamic items = folder.Items;
        items.Sort("[ReceivedTime]", true);
        var emails = new List<object>();
        int total = Math.Min(count, (int)items.Count);
        for (int i = 1; i <= total; i++)
        {
            dynamic item = items[i];
            try
            {
                emails.Add(new { entryId = (string)item.EntryID, subject = (string)item.Subject, sender = (string)item.SenderName, senderEmail = (string)item.SenderEmailAddress, received = ((DateTime)item.ReceivedTime).ToString("yyyy-MM-dd HH:mm:ss"), unread = (bool)item.UnRead, hasAttachments = (int)item.Attachments.Count > 0 });
            }
            catch { }
        }
        var totalCount = (int)folder.Items.Count;
        Release(ns, outlook);
        return new { emails, total = totalCount };
    }

    private static object? GetEmail(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        var entryId = body.GetProperty("entryId").GetString()!;
        dynamic item = ns.GetItemFromID(entryId);

        var attachments = new List<object>();
        for (int i = 1; i <= (int)item.Attachments.Count; i++)
        {
            dynamic att = item.Attachments[i];
            attachments.Add(new { index = i, fileName = (string)att.FileName, size = (long)att.Size });
        }
        var result = new
        {
            entryId,
            subject = (string)item.Subject,
            sender = (string)item.SenderName,
            senderEmail = (string)item.SenderEmailAddress,
            to = (string)item.To,
            cc = (string)item.CC,
            received = ((DateTime)item.ReceivedTime).ToString("yyyy-MM-dd HH:mm:ss"),
            body = (string)item.Body,
            htmlBody = (string)item.HTMLBody,
            unread = (bool)item.UnRead,
            attachments
        };
        Release(ns, outlook);
        return result;
    }

    private static object? SearchEmails(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        var folderPath = body.TryGetProperty("folder", out var fp) ? fp.GetString() : null;
        dynamic folder = GetFolder(ns, folderPath);
        int maxResults = body.TryGetProperty("maxResults", out var mr) ? mr.GetInt32() : 50;

        var filters = new List<string>();
        if (body.TryGetProperty("subject", out var subj)) filters.Add($"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subj.GetString()!}%'");
        if (body.TryGetProperty("sender", out var sender)) filters.Add($"@SQL=\"urn:schemas:httpmail:sendername\" LIKE '%{sender.GetString()!}%'");
        if (body.TryGetProperty("startDate", out var sd)) filters.Add($"[ReceivedTime] >= '{sd.GetString()!}'");
        if (body.TryGetProperty("endDate", out var ed)) filters.Add($"[ReceivedTime] <= '{ed.GetString()!}'");
        if (body.TryGetProperty("unreadOnly", out var uo) && uo.GetBoolean()) filters.Add("[UnRead] = True");

        dynamic items = folder.Items;
        items.Sort("[ReceivedTime]", true);
        if (filters.Count > 0) items = items.Restrict(string.Join(" AND ", filters));

        var emails = new List<object>();
        int total = Math.Min(maxResults, (int)items.Count);
        for (int i = 1; i <= total; i++)
        {
            dynamic item = items[i];
            try { emails.Add(new { entryId = (string)item.EntryID, subject = (string)item.Subject, sender = (string)item.SenderName, received = ((DateTime)item.ReceivedTime).ToString("yyyy-MM-dd HH:mm:ss"), unread = (bool)item.UnRead }); }
            catch { }
        }
        Release(ns, outlook);
        return new { emails, count = emails.Count };
    }

    private static object? MoveEmail(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic item = ns.GetItemFromID(body.GetProperty("entryId").GetString()!);
        dynamic dest = GetFolder(ns, body.GetProperty("targetFolder").GetString()!);
        item.Move(dest);
        Release(ns, outlook);
        return new { message = "Email moved" };
    }

    private static object? DeleteEmail(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic item = ns.GetItemFromID(body.GetProperty("entryId").GetString()!);
        item.Delete();
        Release(ns, outlook);
        return new { message = "Email deleted" };
    }

    private static object? MarkRead(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        bool unread = body.TryGetProperty("unread", out var u) && u.GetBoolean();
        dynamic item = ns.GetItemFromID(body.GetProperty("entryId").GetString()!);
        item.UnRead = unread;
        item.Save();
        Release(ns, outlook);
        return new { message = unread ? "Marked unread" : "Marked read" };
    }

    private static object? Reply(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        bool replyAll = body.TryGetProperty("replyAll", out var ra) && ra.GetBoolean();
        dynamic item = ns.GetItemFromID(body.GetProperty("entryId").GetString()!);
        dynamic reply = replyAll ? item.ReplyAll() : item.Reply();
        reply.Body = body.GetProperty("body").GetString()! + reply.Body;
        bool send = body.TryGetProperty("send", out var s) && s.GetBoolean();
        if (send) reply.Send(); else reply.Display(false);
        Release(ns, outlook);
        return new { message = send ? "Reply sent" : "Reply draft opened" };
    }

    private static object? Forward(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic item = ns.GetItemFromID(body.GetProperty("entryId").GetString()!);
        dynamic fwd = item.Forward();
        fwd.To = body.GetProperty("to").GetString()!;
        if (body.TryGetProperty("body", out var b)) fwd.Body = b.GetString()! + fwd.Body;
        bool send = body.TryGetProperty("send", out var s) && s.GetBoolean();
        if (send) fwd.Send(); else fwd.Display(false);
        Release(ns, outlook);
        return new { message = send ? "Forwarded" : "Forward draft opened" };
    }

    private static object? GetCalendar(string? startDate, string? endDate)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic calendar = ns.GetDefaultFolder(9); // olFolderCalendar
        var start = string.IsNullOrEmpty(startDate) ? DateTime.Today : DateTime.Parse(startDate);
        var end   = string.IsNullOrEmpty(endDate) ? start.AddDays(7) : DateTime.Parse(endDate);

        dynamic items = calendar.Items;
        items.Sort("[Start]");
        items.IncludeRecurrences = true;
        dynamic restricted = items.Restrict($"[Start] >= '{start:M/d/yyyy}' AND [Start] <= '{end:M/d/yyyy 11:59 PM}'");

        var appointments = new List<object>();
        foreach (dynamic item in restricted)
        {
            try { appointments.Add(new { entryId = (string)item.EntryID, subject = (string)item.Subject, start = ((DateTime)item.Start).ToString("yyyy-MM-dd HH:mm"), end = ((DateTime)item.End).ToString("yyyy-MM-dd HH:mm"), location = (string)item.Location, allDay = (bool)item.AllDayEvent, organizer = (string)item.Organizer }); }
            catch { }
        }
        Release(ns, outlook);
        return new { appointments, count = appointments.Count };
    }

    private static object? CreateAppointment(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic appt = outlook.CreateItem(1); // olAppointmentItem
        appt.Subject = body.GetProperty("subject").GetString()!;
        appt.Start = DateTime.Parse(body.GetProperty("start").GetString()!);
        appt.End = DateTime.Parse(body.GetProperty("end").GetString()!);
        if (body.TryGetProperty("location", out var loc)) appt.Location = loc.GetString()!;
        if (body.TryGetProperty("body", out var b)) appt.Body = b.GetString()!;
        if (body.TryGetProperty("allDay", out var ad) && ad.GetBoolean()) appt.AllDayEvent = true;
        if (body.TryGetProperty("attendees", out var atts))
        {
            foreach (var att in atts.EnumerateArray()) appt.Recipients.Add(att.GetString()!);
            appt.MeetingStatus = 1; // olMeeting
        }
        appt.Save();
        bool send = body.TryGetProperty("send", out var s) && s.GetBoolean();
        if (send) appt.Send();
        string entryId = (string)appt.EntryID;
        Release(outlook);
        return new { message = "Appointment created", entryId };
    }

    private static object? DeleteAppointment(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic item = ns.GetItemFromID(body.GetProperty("entryId").GetString()!);
        item.Delete();
        Release(ns, outlook);
        return new { message = "Appointment deleted" };
    }

    private static object? GetContacts(string? folder)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic contactsFolder = string.IsNullOrEmpty(folder) ? ns.GetDefaultFolder(10) : GetFolder(ns, folder);
        var contacts = new List<object>();
        foreach (dynamic item in contactsFolder.Items)
        {
            try { contacts.Add(new { entryId = (string)item.EntryID, fullName = (string)item.FullName, email = (string)item.Email1Address, company = (string)item.CompanyName, phone = (string)item.BusinessTelephoneNumber, jobTitle = (string)item.JobTitle }); }
            catch { }
        }
        Release(ns, outlook);
        return new { contacts, count = contacts.Count };
    }

    private static object? CreateContact(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic contact = outlook.CreateItem(2); // olContactItem
        if (body.TryGetProperty("firstName", out var fn)) contact.FirstName = fn.GetString()!;
        if (body.TryGetProperty("lastName", out var ln)) contact.LastName = ln.GetString()!;
        if (body.TryGetProperty("email", out var em)) contact.Email1Address = em.GetString()!;
        if (body.TryGetProperty("company", out var co)) contact.CompanyName = co.GetString()!;
        if (body.TryGetProperty("phone", out var ph)) contact.BusinessTelephoneNumber = ph.GetString()!;
        if (body.TryGetProperty("jobTitle", out var jt)) contact.JobTitle = jt.GetString()!;
        if (body.TryGetProperty("fullName", out var full)) contact.FullName = full.GetString()!;
        contact.Save();
        string entryId = (string)contact.EntryID;
        Release(outlook);
        return new { message = "Contact created", entryId };
    }

    private static object? SearchContacts(JsonElement body)
    {
        dynamic outlook = GetOutlook();
        dynamic ns = GetNamespace(outlook);
        dynamic folder = ns.GetDefaultFolder(10); // olFolderContacts
        var query = body.GetProperty("query").GetString()!;
        int max = body.TryGetProperty("maxResults", out var mr) ? mr.GetInt32() : 50;

        dynamic items = folder.Items.Restrict($"@SQL=\"urn:schemas:contacts:cn\" LIKE '%{query}%' OR \"urn:schemas:contacts:email1\" LIKE '%{query}%' OR \"urn:schemas:contacts:o\" LIKE '%{query}%'");
        var contacts = new List<object>();
        int count = 0;
        foreach (dynamic item in items)
        {
            if (count >= max) break;
            try { contacts.Add(new { entryId = (string)item.EntryID, fullName = (string)item.FullName, email = (string)item.Email1Address, company = (string)item.CompanyName, phone = (string)item.BusinessTelephoneNumber }); count++; }
            catch { }
        }
        Release(ns, outlook);
        return new { contacts, count = contacts.Count };
    }
}
