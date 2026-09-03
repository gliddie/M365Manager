<#
.SYNOPSIS
    Extracts subject, body text and embedded images from the legacy Outlook .oft templates.
.DESCRIPTION
    The legacy admin scripts didn't send mail themselves - they opened an Outlook template
    (Invoke-Expression -Command <path>.oft) that the operator sent by hand. M365Manager sends
    automatically, so the wording has to be lifted out of those binary templates and into the
    notification builders in Core.

    .oft files are OLE compound documents (MS-CFB) holding MAPI property streams. This reads them
    through Windows' own structured-storage API, so **no Outlook installation is required** - the
    machine this repo is developed on has none, and the "new Outlook" registers no COM automation
    server anyway.

    For each <name>.oft it writes:
      <name>.txt      subject, recipients, sender and the plain-text body
      <name>.html     only when the template carries a PR_HTML stream (most don't - see below)
      <name>.files\   embedded screenshots/logos, plus images.txt mapping each file to its Content-ID

    These templates store their formatted body as compressed RTF (PR_RTF_COMPRESSED) with the HTML
    encapsulated inside it, not as PR_HTML. Decoding that needs an LZFu decompressor plus RTF
    de-encapsulation, which is not implemented here: the plain-text body carries the wording and
    the attachments carry the images, which is what the rewrite actually needs. The Word-generated
    HTML in those templates would have had to be rebuilt by hand regardless.
.EXAMPLE
    pwsh scripts/Export-MailTemplates.ps1
.EXAMPLE
    pwsh scripts/Export-MailTemplates.ps1 -Path D:\templates
#>
[CmdletBinding()]
param(
    # Folder holding the .oft files. Defaults to docs/mail-templates in this repo.
    [string]$Path
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot

Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

// .NET's System.Runtime.InteropServices.ComTypes kept IStream and STATSTG but dropped IStorage and
// IEnumSTATSTG, so both are declared here. Every method must be listed in vtable order even when
// unused - the ones after it would otherwise resolve to the wrong slot.
[ComImport, Guid("0000000D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumSTATSTG
{
    [PreserveSig] int Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] STATSTG[] rgelt, out uint fetched);
    [PreserveSig] int Skip(uint celt);
    void Reset();
    void Clone(out IEnumSTATSTG e);
}

[ComImport, Guid("0000000B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IStorage
{
    void CreateStream([MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint r1, uint r2, out IStream stream);
    void OpenStream([MarshalAs(UnmanagedType.LPWStr)] string name, IntPtr r1, uint mode, uint r2, out IStream stream);
    void CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string name, uint mode, uint r1, uint r2, out IStorage storage);
    void OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage priority, uint mode, IntPtr exclude, uint r, out IStorage storage);
    void CopyTo(uint ciidExclude, [MarshalAs(UnmanagedType.LPArray)] Guid[] rgiidExclude, IntPtr snbExclude, IStorage dest);
    void MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string name, IStorage dest, [MarshalAs(UnmanagedType.LPWStr)] string newName, uint flags);
    void Commit(uint flags);
    void Revert();
    void EnumElements(uint r1, IntPtr r2, uint r3, out IEnumSTATSTG e);
    void DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string name);
    void RenameElement([MarshalAs(UnmanagedType.LPWStr)] string oldName, [MarshalAs(UnmanagedType.LPWStr)] string newName);
    void SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string name, FILETIME ctime, FILETIME atime, FILETIME mtime);
    void SetClass(ref Guid clsid);
    void SetStateBits(uint stateBits, uint mask);
    void Stat(out STATSTG stat, uint flags);
}

/// <summary>Reads the handful of MAPI properties a mail template carries out of a .oft file.</summary>
public static class OftReader
{
    const uint STGM_READ_EXCLUSIVE = 0x00000010; // STGM_READ | STGM_SHARE_EXCLUSIVE

    // MAPI property tags, as the stream names spell them: __substg1.0_<tag><type>.
    // Type 001F = Unicode string, 001E = ANSI string, 0102 = binary.
    const string Subject          = "0037";
    const string NormalizedSubj   = "0E1D";
    const string BodyText         = "1000";
    const string BodyHtml         = "1013";
    const string DisplayTo        = "0E04";
    const string DisplayCc        = "0E03";
    const string SenderName       = "0042";
    const string AttachData       = "3701";
    const string AttachFileName   = "3704";
    const string AttachLongName   = "3707";
    const string AttachMimeTag    = "370E";
    const string AttachContentId  = "3712";

    [DllImport("ole32.dll")]
    static extern int StgOpenStorage(
        [MarshalAs(UnmanagedType.LPWStr)] string name, IStorage priority,
        uint mode, IntPtr exclude, uint reserved, out IStorage storage);

    public sealed class Image
    {
        public string FileName;
        public string ContentId;
        public string MimeType;
        public byte[] Content;
    }

    public sealed class Template
    {
        public string Subject = "";
        public string To = "";
        public string Cc = "";
        public string Sender = "";
        public string BodyText = "";
        public string BodyHtml = "";
        public List<Image> Images = new List<Image>();
    }

    public static Template Read(string path)
    {
        IStorage root;
        int hr = StgOpenStorage(path, null, STGM_READ_EXCLUSIVE, IntPtr.Zero, 0, out root);
        if (hr != 0) throw new COMException("Not a readable compound file: " + path, hr);

        try
        {
            var template = new Template();
            var entries = ListEntries(root);

            template.Subject  = ReadString(root, entries, Subject);
            if (template.Subject.Length == 0) template.Subject = ReadString(root, entries, NormalizedSubj);
            template.To       = ReadString(root, entries, DisplayTo);
            template.Cc       = ReadString(root, entries, DisplayCc);
            template.Sender   = ReadString(root, entries, SenderName);
            template.BodyText = ReadString(root, entries, BodyText);
            template.BodyHtml = ReadString(root, entries, BodyHtml);

            foreach (var entry in entries)
            {
                if (entry.Type != 1 || !entry.Name.StartsWith("__attach")) continue;

                IStorage attachment;
                root.OpenStorage(entry.Name, null, STGM_READ_EXCLUSIVE, IntPtr.Zero, 0, out attachment);
                try { template.Images.Add(ReadAttachment(attachment)); }
                finally { Marshal.ReleaseComObject(attachment); }
            }

            return template;
        }
        finally { Marshal.ReleaseComObject(root); }
    }

    static Image ReadAttachment(IStorage attachment)
    {
        var entries = ListEntries(attachment);
        var name = ReadString(attachment, entries, AttachLongName);
        if (name.Length == 0) name = ReadString(attachment, entries, AttachFileName);

        return new Image
        {
            FileName  = name,
            ContentId = ReadString(attachment, entries, AttachContentId),
            MimeType  = ReadString(attachment, entries, AttachMimeTag),
            Content   = ReadBytes(attachment, entries, AttachData),
        };
    }

    struct Entry { public int Type; public string Name; }

    static List<Entry> ListEntries(IStorage storage)
    {
        var result = new List<Entry>();
        IEnumSTATSTG e;
        storage.EnumElements(0, IntPtr.Zero, 0, out e);
        try
        {
            var stat = new STATSTG[1];
            uint fetched;
            while (e.Next(1, stat, out fetched) == 0 && fetched == 1)
                result.Add(new Entry { Type = stat[0].type, Name = stat[0].pwcsName });
        }
        finally { Marshal.ReleaseComObject(e); }
        return result;
    }

    /// <summary>Finds "__substg1.0_&lt;tag&gt;&lt;type&gt;" for any string/binary type and decodes it.</summary>
    static string ReadString(IStorage storage, List<Entry> entries, string tag)
    {
        foreach (var suffix in new[] { "001F", "001E", "0102" })
        {
            var name = "__substg1.0_" + tag + suffix;
            if (!Contains(entries, name)) continue;

            var bytes = ReadStream(storage, name);
            if (suffix == "001F") return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            if (suffix == "001E") return Encoding.GetEncoding(1252).GetString(bytes).TrimEnd('\0');
            // 0102 on a text property: PR_HTML, stored as raw bytes rather than a string.
            return Encoding.UTF8.GetString(bytes).TrimEnd('\0');
        }
        return "";
    }

    static byte[] ReadBytes(IStorage storage, List<Entry> entries, string tag)
    {
        var name = "__substg1.0_" + tag + "0102";
        return Contains(entries, name) ? ReadStream(storage, name) : new byte[0];
    }

    static bool Contains(List<Entry> entries, string name)
    {
        foreach (var entry in entries)
            if (entry.Type == 2 && string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    static byte[] ReadStream(IStorage storage, string name)
    {
        IStream stream;
        storage.OpenStream(name, IntPtr.Zero, STGM_READ_EXCLUSIVE, 0, out stream);
        try
        {
            STATSTG stat;
            stream.Stat(out stat, 1);
            var buffer = new byte[(int)stat.cbSize];
            IntPtr read = Marshal.AllocHGlobal(IntPtr.Size);
            try { stream.Read(buffer, buffer.Length, read); return buffer; }
            finally { Marshal.FreeHGlobal(read); }
        }
        finally { Marshal.ReleaseComObject(stream); }
    }
}
'@

if (-not $Path) { $Path = Join-Path $repoRoot "docs\mail-templates" }
if (-not (Test-Path $Path)) { throw "Folder not found: $Path" }

$templates = Get-ChildItem -Path $Path -Filter *.oft -File
if ($templates.Count -eq 0) { throw "No .oft files in $Path" }

foreach ($file in $templates) {
    Write-Host "Reading $($file.Name) ..."
    $t = [OftReader]::Read($file.FullName)

    $meta = [Collections.Generic.List[string]]::new()
    $meta.Add("Subject: $($t.Subject)")
    $meta.Add("To: $($t.To)")
    $meta.Add("CC: $($t.Cc)")
    $meta.Add("Sender: $($t.Sender)")
    $meta.Add("")
    $meta.Add("----- body (plain text) -----")
    $meta.Add($t.BodyText)
    Set-Content -Path ([IO.Path]::ChangeExtension($file.FullName, ".txt")) -Value $meta -Encoding UTF8

    if ($t.BodyHtml) {
        Set-Content -Path ([IO.Path]::ChangeExtension($file.FullName, ".html")) -Value $t.BodyHtml -Encoding UTF8
    }

    if ($t.Images.Count -gt 0) {
        $assetDir = [IO.Path]::ChangeExtension($file.FullName, ".files")
        New-Item -ItemType Directory -Path $assetDir -Force | Out-Null

        $index = [Collections.Generic.List[string]]::new()
        $n = 0
        foreach ($image in $t.Images) {
            $n++
            $name = if ($image.FileName) { $image.FileName } else { "image$n.bin" }
            [IO.File]::WriteAllBytes((Join-Path $assetDir $name), $image.Content)
            $kb = [math]::Round($image.Content.Length / 1KB, 1)
            $index.Add("$name`t${kb} KB`t$($image.MimeType)`tcid:$($image.ContentId)")
        }
        Set-Content -Path (Join-Path $assetDir "images.txt") -Value $index -Encoding UTF8
    }

    Write-Host "  subject: $($t.Subject)"
    Write-Host "  body: $($t.BodyText.Length) chars, html: $($t.BodyHtml.Length) chars, images: $($t.Images.Count)"
}

Write-Host "Done. $($templates.Count) template(s) exported to $Path"
