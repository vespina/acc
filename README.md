# ACC
Active Connections Controller

#### Objective
This class allows to control the number of available connections in a shared app. Can be used to restrict the amount of users or workstations running an specific app.

#### Usage
Each user most create and setup an instance of this class, and then try to establish a connection:

    SET PROCEDURE TO acc ADDITIVE
    PUBLIC goACC
    goACC = CREATEOBJECT("VEActiveConnectionsController")
    WITH goACC
        .sharedFolder = ".\TEMP"
        .maxConnections = 10
    ENDWITH
    
    IF goACC.Connect() > 0
      * YOU ARE GOOD TO GO
    ELSE 
      * NO MORE AVAILABLE CONNECTIONS OR AN ERROR OCURRED
      MESSAGEBOX(goACC.lastError)
      RETURN
    ENDIF

    ......
    
    goACC.Disconnect()
    
    
Altough is recommended to manually call the Disconnect() method in the app closing procedure, the connection will be automatically released upon the app is closed, even if Disconnect method is not called.


