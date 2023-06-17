# VBJSON
🔎 Visual Basic JSON parser / encoder (fork)

> VBJSON is a VB6 adaptation of the VBA JSON project at http://code.google.com/p/vba-json/
> Some bugs fixed, speed improvements added for VB6 by Michael Glaser (vbjson@ediy.co.nz)
> BSD Licensed

### What's Changed

- Migrate to Class Module
- Prefer `System.Text.StringBuilder` over vbAccelerator's pure VB String Builder implementation for faster speed

### Usage:

```visual basic
Dim JSON As New cJSON
Dim Students As Collection
Set Students = JSON.Parse("[{""Name"": {""FirstName"": ""Yidaozhan"", ""LastName"": ""Ya""}, ""Age"": 16}]")
MsgBox Students(1)!Name!FirstName  'Yidaozhan
```
