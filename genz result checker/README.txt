Use only these files on the frontend:
- index.html
- admin.html
- student.html

Apps Script:
- Code.gs only

Steps:
1. Replace your current Code.gs with this new Code.gs
2. Save
3. Run setupSystem() once
4. Deploy / Update the Apps Script Web App
   - Execute as: Me
   - Who has access: Anyone
5. On GitHub/Vercel, upload only:
   - index.html
   - admin.html
   - student.html

Notes:
- This build uses your Apps Script WEB_APP_URL directly through JSONP, so no extra frontend files are needed.
- Passport link is handled in admin only.
- Student phone field is removed.
- Category is now a dropdown.
