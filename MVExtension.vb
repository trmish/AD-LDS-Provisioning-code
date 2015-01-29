Imports Microsoft.MetadirectoryServices

Public Class MVExtensionObject
    Implements IMVSynchronization

    Public Sub Initialize() Implements IMvSynchronization.Initialize
        ' TODO: Add initialization code here
    End Sub

    Public Sub Terminate() Implements IMvSynchronization.Terminate
        ' TODO: Add termination code here
    End Sub

    Public Sub Provision(ByVal mventry As MVEntry) Implements IMVSynchronization.Provision
        Dim container As String
        Dim rdn As String
        Dim ADMA As ConnectedMA
        Dim numConnectors As Integer
        Dim myConnector As CSEntry
        Dim csentry As CSEntry
        Dim dn As ReferenceValue

        If (mventry.ObjectType.Equals("person") Or mventry.ObjectType.Equals("user") Or mventry.ObjectType.Equals("group")) Then
            ' Ensure that the cn attribute is present.
            If Not mventry("cn").IsPresent Then
                Throw New UnexpectedDataException("cn attribute is not present.")
            End If

            Try
                ' Calculate the container and RDN.
                container = "cn=Users,cn=partition1,DC=lds,DC=local"
                rdn = "cn=" & mventry("cn").Value
                'Throw New UnexpectedDataException(rdn)
                ADMA = mventry.ConnectedMAs("ADLDS")
                dn = ADMA.EscapeDNComponent(rdn).Concat(container)
                ' Throw New UnexpectedDataException(dn.ToString)
                numConnectors = ADMA.Connectors.Count
                'Logging.Log("!!! start", True, 2)
                'Logging.Log("!!! start", True, 0)
                'Logging.Log("!!! --- " & dn.ToString, True, 2)
                ' create a new connector.
                If numConnectors = 0 Then
                    If (mventry.ObjectType.Equals("person") Or mventry.ObjectType.Equals("user")) Then
                        csentry = ADMA.Connectors.StartNewConnector("user")
                        csentry.DN = dn
                        csentry("unicodePwd").Value = "Password1"
                        csentry.CommitNewConnector()
                        'Logging.Log("!!! add " & dn.ToString, True, 2)
                    End If
                    If (mventry.ObjectType.Equals("group")) Then
                        csentry = ADMA.Connectors.StartNewConnector("group")
                        csentry.DN = dn
                        csentry.CommitNewConnector()
                        'Logging.Log("!!! ggg " & dn.ToString, True, 2)
                    End If
                ElseIf numConnectors = 1 Then
                    ' If the connector has a different DN rename it.
                    myConnector = ADMA.Connectors.ByIndex(0)
                    myConnector.DN = dn
                    'Logging.Log("!!! rename " & dn.ToString, True, 2)
                Else
                    Throw New UnexpectedDataException("Error: There are" + numConnectors.ToString + " connectors")
                End If
            Catch ex As ObjectAlreadyExistsException
                'Logging.Log("!!! ObjectAlreadyExistsException " & dn.ToString, True, 2)
            Catch ex As Exception
                'Logging.Log("!!! Exception " & dn.ToString, True, 2)
            End Try
        End If
    End Sub

    Public Function ShouldDeleteFromMV(ByVal csentry As CSEntry, ByVal mventry As MVEntry) As Boolean Implements IMVSynchronization.ShouldDeleteFromMV
        ' TODO: Add MV deletion code here
        Throw New EntryPointNotImplementedException()
    End Function
End Class
