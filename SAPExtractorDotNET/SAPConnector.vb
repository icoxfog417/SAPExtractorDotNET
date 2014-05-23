Imports Microsoft.VisualBasic
Imports SAP.Middleware.Connector

Namespace SAPExtractorDotNET

    ''' <summary>
    ''' Create Connection to SAP.<br/>
    ''' implements Serializable to store in session.
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()>
    Public Class SAPConnector

        Private _destination As String
        ''' <summary>destination name defined in web.config/app.config</summary>
        Public ReadOnly Property Destination As String
            Get
                Return _destination
            End Get
        End Property

        Private _isSilent As Boolean = False
        ''' <summary>is silent Login or not (silent mean user/password is already defined by destination)</summary>
        Public ReadOnly Property IsSilent As Boolean
            Get
                Return _isSilent
            End Get
        End Property

        Private _user As String = ""
        Public ReadOnly Property User() As String
            Get
                Return _user
            End Get
        End Property

        Private _password As String = ""
        Private Property Password() As String
            Get
                Dim bytePass As Byte() = Convert.FromBase64String(_password)
                Return System.Text.Encoding.Default.GetString(bytePass)
            End Get
            Set(value As String)
                Dim byted() As Byte = System.Text.Encoding.Default.GetBytes(value)
                _password = Convert.ToBase64String(byted)
            End Set
        End Property

        ''' <summary>
        ''' Create Connection only by RFC definition in web.config/app.config
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Login() As RfcDestination
            Dim connection As RfcDestination = Nothing
            Dim destConnection As RfcDestination = RfcDestinationManager.GetDestination(Destination)
            If IsSilent Then
                connection = destConnection
            Else
                Dim userLogin As RfcCustomDestination = destConnection.CreateCustomDestination
                userLogin.User = User
                userLogin.Password = Password
                connection = userLogin
            End If

            If connection IsNot Nothing Then
                connection.Ping() 'confirm connection
            End If

            Return connection

        End Function

        Public Sub New(ByVal destination As String)
            _destination = destination
            _isSilent = True
        End Sub

        Public Sub New(ByVal destination As String, ByVal user As String, ByVal password As String)
            _destination = destination
            _isSilent = False
            _user = user
            Me.Password = password
        End Sub

    End Class

End Namespace
