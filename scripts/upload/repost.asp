<%

' Collect Stats from POST operation as redirect as GET

Response.Buffer = TRUE
FileName = Request.Form("FileName")
FileExtention = Request.Form("FileExtention")
FileSize = Request.Form("FileSize")
FilePath = Request.Form("FilePath")
TranID = Request.Form("TranID")
GalleryID = Request.Form("GalleryID")
ContactID = Request.Form("ContactID")
ImageTitle = Request.Form("ImageTitle")
CameraPerson = Request.Form("CameraPerson")

Response.Redirect("upstat.asp?TranID=" & TranID & "&FileName=" & FileName & "&FileSize=" & FileSize & "&FileExtention=" & FileExtention & "&GalleryID=" & GalleryID & "&ContactID=" & ContactID & "&CameraPerson=" & CameraPerson & "&ImageTitle=" & ImageTitle)

%>
