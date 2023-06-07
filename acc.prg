* ACC (Active Connections Controller)
* Controlador de conexiones concurrentes
*
* Esa clase permite determinar el nro. de
* estaciones que estan ejecutando un programa
* cualquiera, y a la vez permite registrar
* la ejecución del programa actual
*
* Autor: Victor Espina
* Fecha: Noviembre 2006
*
* El presente código puede ser usado a 
* conveniencia del usuario. El autor
* no es responsable de cualquier daño 
* ocasionado por el uso de este código.
*
* Modo de uso:
*
* Al inicio del programa principal, colocar:
*
* PUBLIC goACC
* goACC = CREATEOBJECT("VEActiveConnectionsController")
* goACC.SharedFolder = ".\ACC"
* goACC.MaxConnections = 10  
*
* IF NOT goACC.Connect()
*  MESSAGEBOX(goACC.LastError)
*  QUIT
* ENDIF
*
* 
* Al final del programa principal:
*
* goACC.Disconnect()
*
*
* Marzo 2, 2012 - Victor Espina
* -----------------------------
* Se anexaron dos propiedades nuevas y un metodo a la clase, para permitir la validacion
* automatica de la validez de la conexion establecida:
*
* checkConnectionEvery: 
*  Permite indicar cada cuantos minutos la clase debera verificar si el archivo de marca
*  creado sigue siendo valido. Si se indica cero (0) no se realizara la verificacion.
*
* onConnectionLost:
*  Comando a ejecutar si el archivo de marca deja de ser valido. En VFP6 debe ser una 
*  sentencia simple, pero de VFP7 puede ser un codigo complejo.
*
* IsAlive()
*  Metodo que permite verificar si el archivo de marca aun es valido o no. 
*
DEFINE CLASS VEActiveConnectionsController AS Custom
 *
 SharedFolder = ".\"  && Ubicacion de la carpeta compartida a utilizar para crear los archivos de marca
 MaxConnections = 0   && Nro. máximo de conexiones permitidas
 WorkstationID = ""   && ID de la estacion. Si no se indica, se asume el nombre del equipo.
 MarkFileExt = "ACM"  && Extension de los archivos de marca. Si no se indica se asume .ACM
 LastError = "" 	  && Texto del ultimo error ocurrido
 checkConnectionEvery = 0  && Frecuencia (en minutos) para la verificacion del archivo de marca (0 = nunca)
 onConnectionLost = ""     && Codigo a ejecutar si la conexion con el archivo de marca se pierde
 HIDDEN nFH			  && Handle del archivo de marca correspondiente al proceso actual
 HIDDEN oTimer1       && Timer para verificacion de estado de conexion
 
 
 * Class constructor
 * Constructor de la clase
 *
 PROC Init()
  *
  THIS.WorkstationID = ALLT(LEFT(SYS(0),AT("#",SYS(0)) - 1))
  THIS.nFH = 0
  THIS.oTimer1 = CREATE("VEACCTimer")
  THIS.oTimer1.Enabled = .F.
  *
 ENDPROC


 * GetCurrentMarkFile 
 * Devuelve el nombre y ruta del archivo de marca correspondiente
 * a la estacion actual
 *
 PROC GetCurrentMarkFile()
  *
  LOCAL cMarkFile
  cMarkFile=FORCEEXT(THIS.WorkstationID,THIS.MarkFileExt)
  cMarkFile=FORCEPATH(cMarkFile,THIS.SharedFolder)
  cMarkFile=LOWER(cMarkFile)
  
  RETURN cMarkFile
  *
 ENDPROC
 
 
 * GetActiveConnectionsCount
 * Devuelve el nro. de conexiones concurrentes activas. Esto se logra
 * contanto cuantos archivos existentes en la carpeta compartida aun
 * estan bloqueados por otro proceso.
 *
 PROC GetActiveConnectionsCount()
  *
  LOCAL nActiveCount,nCount,i,cFile,nFH
  LOCAL ARRAY aFiles[1]
  nCount=ADIR(aFiles,ADDBS(THIS.SharedFolder)+"*."+THIS.MarkFileExt)
  nActiveCount = 0
  
  FOR i=1 TO nCount
   *
   * Se obtiene el nombre y ubicacion del archivo de marca a validar
   cFile=LOWER(FORCEPATH(aFiles[i,1],THIS.SharedFolder))
   
   * Se intenta abrir el archivo de marca para escritura
   nFH=FOPEN(cFile,1)
    
   * Si no se pudo abrir el archivo significa que hay un proceso activo
   * que aun lo tiene bloqueado, por lo que se cuenta como una conexion
   * activa, de lo contrario se cierra el archivo y se elimina pues 
   * corresponde a una conexion que termino anormalmente (ya que si 
   * hubiera terminado normalmente, el archivo habria sido borrado por
   * la aplicacion directamente).
   IF nFH < 0 
    nActiveCount=nActiveCount + 1
   ELSE
    FCLOSE(nFH)
    ERASE (cFile)
   ENDIF
   *
  ENDFOR
  
  
  RETURN nActiveCount
  *
 ENDPROC
 
 
 * Connect
 * Determina si hay conexiones disponibles y procede a crear
 * un archivo de marca. El metodo devuelte:
 *
 * 1  si se pudo crear la conexion
 * 0  si no hay conexiones disponibles
 * -1 la estacion ya esta conectada
 * -2 si no se pudo crear el archivo de marca
 *
 PROC Connect()
  *
  * Se determina la cantidad de conexiones activas
  LOCAL nActiveCount
  nActiveCount = THIS.GetActiveConnectionsCount()
  
  * Si no hay mas conexiones disponibles, se cancela
  * en este punto. Se utiliza >= y no solo = por razones
  * de programacion defensiva.
  IF nActiveCount >= THIS.MaxConnections
   THIS.LastError = "No hay conexiones disponibles"
   RETURN 0
  ENDIF
  
  * Si ya exite un archivo de marca para la estacion, se
  * cancela pues se asume que el programa ya esta en
  * ejecucion en la estacion
  LOCAL cMarkFile
  cMarkFile = THIS.GetCurrentMarkFile()
  IF FILE(cMarkFile)
   THIS.LastError = "Esta estación ya está conectada"
   RETURN -1
  ENDIF
  
  
  * Se crea el de marca
  THIS.nFH = FCREATE(cMarkFile)
  IF THIS.nFH < 0
   THIS.LastError = "No se pudo crear el archivo " + cMarkFile
   RETURN -2
  ENDIF
  
  * Si se indico un intervalo para verificar la conexion, se configura el timer y se inicia
  IF THIS.checkConnectionEvery > 0
   THIS.oTimer1.Set(THIS)
  ENDIF

  RETURN 1  
  *
 ENDPROC


 
 * Disconnect
 * Libera el archivo de marca correspondiente al proceso actual
 *
 PROC Disconnect()
  *
  * Si no hay un archivo de marca creado, se cancela
  IF THIS.nFH = 0
   RETURN
  ENDIF
  
  * Se cierra y elimina el archivo de marca
  LOCAL cMarkFile
  cMarkFile = THIS.GetCurrentMarkFile()
  FCLOSE(THIS.nFH)
  ERASE (cMarkFile)
  
  * Se libera el timer de verificacion
  THIS.oTimer1.Clear()
  *
 ENDPROC
 
 PROC foo
  FCLOSE(THIS.nFH)
 ENDPROC
 
 * IsAlive
 * Determina si el archivo de marca aun es valido
 *
 PROC IsAlive
  *
  * Si no hay un archivo de marca creado, se cancela
  IF THIS.nFH = 0
   RETURN .F.
  ENDIF
  
  RETURN FFLUSH(THIS.nFH)
  *
 ENDPROC
 *
ENDDEFINE


* VEACCTimer
* Timer de verificacion de conexion para VEActiveConnectionController
*
DEFINE CLASS VEACCTimer AS Timer
 *
 Target = NULL
 
 PROCEDURE Set(poTarget)
  THIS.Target = poTarget
  THIS.Interval = poTarget.checkConnectionEvery * 60 * 1000
  THIS.Enabled = .T.
 ENDPROC
 
 PROCEDURE Timer
  THIS.Enabled = .F.
  ?"Paso"
  IF THIS.Target.IsAlive()
   THIS.Enabled = .T.
   RETURN
  ENDIF
  IF !EMPTY(THIS.Target.onConnectionLost)
   LOCAL cCmd
   cCmd = THIS.Target.onConnectionLost
   IF " 06.00" $ VERSION()
    &cCmd
   ELSE
    EXECSCRIPT(cCmd)
   ENDIF
  ENDIF
 ENDPROC
 
 PROCEDURE Clear
  THIS.Enabled = .F.
  THIS.Target = NULL  
 ENDPROC
 *
ENDDEFINE



