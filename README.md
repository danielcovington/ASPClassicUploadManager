# UploadManager – Classic ASP “PHP‑style” File Uploads

A small **C# / .NET Framework 4.8** COM‑visible DLL that parses `multipart/form‑data`
posts so Classic ASP can work with file uploads the same way PHP’s `$_FILES` does.

---

## Table of Contents
1. [Features](#features)  
2. [Prerequisites](#prerequisites)  
3. [Building the DLL](#building-the-dll)  
4. [Registering for COM](#registering-for-com)  
5. [IIS / Classic ASP Size Limits](#iis--classicasp-size-limits)  
6. [Using in ASP Classic](#using-in-asp-classic)  
7. [Error Handling](#error-handling)  
8. [Troubleshooting](#troubleshooting)

---

## Features
- Parses `multipart/form‑data` **in‑memory** – no temporary files needed.  
- Exposes each upload as an **`UploadedFile`** object with  
  `FieldName`, `FileName`, `ContentType`, `Size`, and `SaveAs(path)`.  
- Enumerable from VBScript (`For Each f In uploadMgr`).  
- Throws **`COMException`** with helpful messages that surface to  
  `Err.Number / Err.Description`.  
- Tested on IIS 10 / Windows Server 2022, .NET Fx 4.8.

---

## Prerequisites
| Requirement | Version |
|-------------|---------|
| Windows | 10 / Server 2016 + |
| Visual Studio | 2022 / 2025 (Community or higher) |
| .NET Framework Dev Pack | **4.8** |
| Privileges | Local admin (for `regasm`, IIS edits) |

---

## Building the DLL

1. **Create the project**

   ```text
   File ▸ New ▸ Project… ▸ “Class Library (.NET Framework)”
   Target framework: .NET Framework 4.8
   Project name:     UploadManager
   ```

2. **Make the assembly COM‑visible**

   *Project ▸ Properties ▸ Application ▸ **Assembly Information…***  
   - ☑️ **Make assembly COM‑visible**  
   - *Signing tab* → ☑️ **Sign the assembly** → create a key.

3. **Add code files**

   ```text
   Delete Class1.cs
   Add ▸ Class…  →  UploadedFile.cs  (paste code)
   Add ▸ Class…  →  UploadManager.cs (paste code)
   ```

4. **Add a library‑wide GUID** (`Properties/AssemblyInfo.cs`)

   ```csharp
   [assembly: Guid("5A23B644-6F4F-4BBC-9CE4-D3325C8C87F1")]
   ```

5. **(Optional) auto‑register on build** – edit `.csproj`

   ```xml
   <PropertyGroup>
     <RegisterAssembly>true</RegisterAssembly>
     <RegisterAssemblyMSBuildArchitecture>x64</RegisterAssemblyMSBuildArchitecture>
   </PropertyGroup>
   ```

6. **Build**

   ```text
   Configuration: Release
   Platform:      Any CPU
   Build ▸ Build Solution
   ```

   *“Registration succeeded”* appears if auto‑register is on.

---

## Registering for COM (manual)

```powershell
cd .\UploadManager\bin\Release
regasm UploadManager.dll /codebase      # adds CLSID + path to registry
gacutil -i UploadManager.dll            # optional: add to GAC
```

> **Tip:** Stop IIS / recycle the App Pool before rebuilding; otherwise the DLL
> may be locked.

---

## IIS / Classic ASP Size Limits

Classic ASP default limit ≈ 200 KB.  
Raise limits in **web.config** (example = 50 MB):

```xml
<configuration>
  <system.webServer>

    <!-- Classic ASP engine -->
    <asp>
      <limits maxRequestEntityAllowed="52428800" />
    </asp>

    <!-- IIS request filtering -->
    <security>
      <requestFiltering>
        <requestLimits maxAllowedContentLength="52428800" />
      </requestFiltering>
    </security>

    <!-- Optional (ARR / reverse proxy) -->
    <!-- <serverRuntime uploadReadAheadSize="52428800" /> -->

  </system.webServer>
</configuration>
```

Recycle the App Pool or run:

```powershell
iisreset /restart
```

---

## Using in ASP Classic

### HTML Form

```html
<form action="/upload.asp" method="post" enctype="multipart/form-data">
  <input type="file" name="dbfile">
  <button>Upload</button>
</form>
```

### `upload.asp`

```asp
<%
Option Explicit
On Error Resume Next

' 1. Read entire body
Dim total : total = Request.TotalBytes
If total = 0 Then
  Response.Write "Upload rejected (check IIS size limits)."
  Response.End
End If

Dim body  : body  = Request.BinaryRead(total)
Dim ctype : ctype = Request.ServerVariables("CONTENT_TYPE")

' 2. Parse with COM object
Dim up : Set up = Server.CreateObject("UploadManagerLib.UploadManager")
up.Parse body, ctype

If Err.Number <> 0 Then
  Response.Write "Error 0x" & Hex(Err.Number) & ": " & Err.Description
  Response.End
End If

' 3. Process files
Dim f, saveDir : saveDir = Server.MapPath("/uploads/")
For Each f In up
  f.SaveAs saveDir & "\" & f.FileName
  Response.Write "Saved " & f.FileName & " (" & f.Size & " bytes)<br>"
Next
%>
```

---

## Error Handling

* **C#** re‑throws all issues as `COMException`.  
* **VBScript** traps them with:

  ```asp
  On Error Resume Next
  up.Parse body, ctype
  If Err.Number <> 0 Then ...  ' Err.Description holds message
  ```

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| `ActiveX component can’t create object` | DLL not registered, or 32‑bit reg + 64‑bit IIS. Re‑register matching architecture. |
| `FileCount = 0`, `TotalBytes = 0` | Increase `maxRequestEntityAllowed` & `maxAllowedContentLength`. |
| Build error “file in use” | Recycle App Pool / stop IIS before rebuild, or turn off auto‑register. |
| `Request is not multipart/form-data` | Ensure form uses `POST` & `enctype="multipart/form-data"`. |

---

**Enjoy hassle‑free file uploads in Classic ASP!**
