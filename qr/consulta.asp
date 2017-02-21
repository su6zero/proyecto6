<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' me fijo si viene algun valor en el querystring, si no viene nada, no hago nada
if request.querystring("emailUsuario") <> "" then
    email = request.querystring("emailUsuario")
    if email = "1@1.com" then
       vari="Si, existe"
       vari=vari & "Esto es una prueba de la capacidad de AJAX"
       vari=vari & "<table><tr><th>Esperemos que funcione como dice</th></tr><tr><td>Esperemos que funcione correctamente en el cuadrante 1,2</td></tr></table>"

    else
        vari="hola compañero"
       vari=vari & "No existe"&vari&"-"&vari&"-"&vari&"-"&vari&"-"&vari&"-"&vari&"-"&vari&"-"
    end if
    response.Write vari
end if
%>
