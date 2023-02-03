# FLOW OAUTH FOR VBA

Utilice la clase Flow Outh FOR VBA para crear una instancia que pueda autenticarse y autorizar  a su aplicación, para poder consumir las Apis de Google (Api Google Drive-Api Google Sheets).

Esta clase ha sido desarrollada en VBA,por lo que encontrará los archivos con los códigos fuente en el directorio de* **"lib me"*** de este repositorio. El libro de trabajo de Excel **"*Oauth.xlsm*"** se encuentra equipado con todos los módulos y clases necesarias para su funcionamiento, ademas de mostrar un ejemplo en módulo ***"Main"***. Sí bien es sierto este caso se muestra en Excel pero no se límita a el, puede usarlo en otras aplicaciones Office activando las referencias correspondientes y descargando la libreria de terceros que se detallan líneas  abajo.

## ¿Cómo funciona Flow Oauth?

La siguiente gráfica muestra de forma abreviada como actua el flujo de Oauth:

[![OAuth 2.0 para aplicaciones de servidor web](https://developers.google.com/static/identity/protocols/oauth2/images/flows/authorization-code.png?hl=es-419 "OAuth 2.0 para aplicaciones de servidor web")](https://developers.google.com/static/identity/protocols/oauth2/images/flows/authorization-code.png?hl=es-419 "OAuth 2.0 para aplicaciones de servidor web")

Si deseas obtener más detalles,consulta [Usa OAuth 2.0 para aplicaciones de servidor web.](https://developers.google.com/identity/protocols/oauth2/web-server?hl=es-419&utm_source=devtools "Usa OAuth 2.0 para aplicaciones de servidor web.")

## Referencias a habilitar
Debes asegurarte tener habilitadas la siguientes referencias antes de comenzar:
- **Microsoft Office 16.0 Object Library**
- **Microsoft Scripting Runtime**
- **Microsoft XML,v6.0**

## Biblioteca de terceros
>Gran parte de nuestro trabajo no hubiera sido posible sin la ayuda del siguiente módulo que nos ayuda a PARSEAR la respuesta del servidor en formato JSON.

Descarga e instala el siguiente módulo del respositorio [https://github.com/VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON "https://github.com/VBA-tools/VBA-JSON")

- ** JsonConverter.bas v2.3.1 **

## Crea un proyecto en Google Api Console

Para poder consumir las Apis de Google es necesario crear un proyecto en ***Google Cloud Platform*** vea [Cómo usar OAuth 2.0 para acceder a las API de Google](https://developers.google.com/identity/protocols/oauth2?hl=es-419 "Cómo usar OAuth 2.0 para acceder a las API de Google") . Luego de haber activado las referencias y descargado el módulo correspondiente.

Deberas crear las siguientes credenciales:

- **ID de cliente de Oauth**
- **Clave de Api**

Guarda estas credenciales en un lugar seguro.

Vea el siguiente video para mayor detalle de como crear las credenciales pulsando [Crear proyecto parte 1](https://www.youtube.com/watch?v=8GG7LnaMtuE&list=PLebWFysFNi3AuZOqFzKNzqHc6mPkkz1AX&index=10 "Crear proyecto parte 1")

# Por fin ...... un ejemplo

Después de la travesía, por fin podemos ver como funciona:

La instancia de FlowOauth deberá comenzar con el método **Initilize** el cual deberá recibir 3 argumentos, todos ellos con las rutas hacia las credenciales que se crearon en [Google Api Console](https://console.developers.google.com/?hl=es-419 "Google Api Console "Google Api Console"), excepto por el **token** que la primera vez que se ejecute este procedimiento se no se tendrá el archivo con dicha información, pero si que debemos pasarle la ruta de donde se alojará dicho archivo en formato JSON.

### ¿Y qué hará?

- Nuestra instancia se encargará de crear un archivo con los siguientes valores **access_token, refresh_token, expires_in, token_type, scope** en la ruta que se le hayamos pasado como argumento en **credentialsToken** para lo cual nos mostrará un pantalla de consentimiento donde debemos aceptar los permisos  correspondientes esto será la primera vez, seguido se nos mostrará un cuadro de dialogo por parte de VBA para ingresar el valor del **code** mostrado en la URL. 
Debería tener un archivo como el siguiente:

```json
{
  "access_token": "1/fFAGRNJru1FTz70BzhT3Zg",
  "expires_in": 3920,
  "token_type": "Bearer",
  "scope": "https://www.googleapis.com/auth/drive.metadata.readonly",
  "refresh_token": "1//xEoDL4iW3cxlI7yDbSRFYNG01kVKM2C-259HOF2aQbI"
}
```


- En caso que se necesite actualizar el token después de haber alcanzado el tiempo de validez, la misma instancia se encargará de solicitar un nuevo **access_token**  utilizando el **refresh_token** .

- En caso que se haya revocado un token o deje  de funcionar pulse [Aquí](https://developers.google.com/identity/protocols/oauth2?hl=es-419#expiration "Aquí") para ver los motivos, se le mostrará un mensaje indicando que se necesita eliminar el token creado inicialmente para solicitar un nuevo, con lo cual empezerá el flujo de nuevo.


```vb
Sub test_Oa()
	
    Dim Ou As New FlowOauth
    Dim client As String
    Dim token As String
    Dim apiKey As String
 
    apiKey = ThisWorkbook.Path & "\credentials\api_key.json"
    token = ThisWorkbook.Path & "\credentials\token.json"
    client = ThisWorkbook.Path & "\credentials\client_secret.json"
    
    With Ou
        .InitializeFlow client, token, apiKey, OU_SCOPE_DRIVE_READONLY
        Debug.Print "API KEY"; " -- "; .GetApiKey
        Debug.Print "TOKEN ACCESS "; " -- "; .GetTokenAccess
    End With

End Sub
```
Sí todo el flujo ha ido bien debería ver en su ventana inmediato un mensaje como el que se muestra acontinuación:

    Flujo de Oauth 2.0 2/02/2023 15:42:25  >>> flow started
    Flujo de Oauth 2.0 2/02/2023 15:42:25  >>> FILE NOT FOUND D:\Mis documentos personales\Proyectos en VBA\Flujo deOauth2.0\credentials\token.json
    Flujo de Oauth 2.0 2/02/2023 15:42:51  >>> new token generated
    API KEY -- AIzaSyAvOlAOE***********************************************************************
    TOKEN ACCESS  -- ya29.a0AVvZVsqA0wR4oxCM2dTrHml***************************************
    

### ¿qué tenemos?

```vb
'Comenzando
.InitializeFlow (credentialsClient,credentialsToken,credentialsApiKey,scope)
	'param : credentialsClient = debe indicar la ruta del archivo generado en Id de cliente de Oauth
	'param: credentialsToken = la ruta de donde se creará el token o donde de se encuentra.
	'param: credentialsApiKey = ruta donde se encuentra el archivo json con el api key.
	'parame: scope = una o mas alcances proporcionados en constantes junto con este repositorio.
	'return: empty


'Revocando un token
.revokeToken(credentialsToken) as boolean
	'param->credentialsToken ->string->ruta donde se encuentra el token
	'return->boolean

'Cambiando de navegador
.webBrowser=[escritura | lectura]
	'value->string->asigne un navegador.
	'Use esta propiedad para cambiar el navegador donde se mostrara la pantalla de consentimiento, por defecto usa Chrome.exe, debe tener el navegador incluido en el PATH o simplemente pasar la ruta completa.

.Operation=[lectura]
	'value->integer->Internamente se utiliza peticiones HTTP para crear y actualizar el token, use esta propiedad para ver el status de la petición HTTP.



```



