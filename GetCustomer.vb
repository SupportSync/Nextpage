Imports System.Net
Imports System.Net.Http
Imports System.Web.Http
Imports Microsoft.AspNet.Identity
Imports System.Web.Http.Description
Imports System.Configuration.ConfigurationManager
Imports SupportSyncCRM.SupportSync.Security
Imports SupportSyncCRM.SupportSync.CustomHelpPageAttributes
Imports System.ComponentModel.DataAnnotations
Imports System.Web.Mvc
Imports System.Web.Http.Cors

Partial Public Class CustomersController
    Inherits ApiController

    ' GET api/<app>/<controller>/<id>
    ''' <summary>
    ''' Gets a customer or a list of customers based on email address 
    ''' </summary>
    ''' <param name="app">Name of application</param>
    ''' <param name="id">Base64-encoded comma-separated list of email addresses</param>
    ''' <returns>Returns an array of customers including shipping address</returns>
    ''' <remarks>
    ''' </remarks>
    <ApiAuthorizeAttribute>
    Public Function GetCustomer(ByVal app As String, ByVal id As String) As List(Of Dynamic.ExpandoObject)

        theCustomerManager = New CustomerManager(UserManager.GetUserId, OrganizationManager.GetOrganizationId)
        theCustomer = New Customer

        Dim theCustomerList As New List(Of Dynamic.ExpandoObject)

        If Helpers.IsBase64Encoded(id) Then

            id = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(id))

            For Each emailAddress As String In id.Split(",")
                If Regex.IsMatch(emailAddress.Trim, AppSettings("RegexEmail")) Then
                    theCustomer = theCustomerManager.GetCustomerByEmail(emailAddress.Trim)

                    If Not theCustomer Is Nothing Then

                        'Find out if customer exists for current Brand
                        If theCustomerManager.IsCustomerInBrand(emailAddress.Trim, BrandManager.GetBrandId) Then

                            theCustomer = theCustomerManager.GetCustomer(theCustomer.CustomerUserId)

                            If Not theCustomer Is Nothing Then
                                Dim obj = CustomerUtility.ConvertToFlatJson(theCustomer)
                                theCustomerList.Add(obj)
                                theCustomer = Nothing
                            End If

                        End If

                    End If

                End If
            Next

        End If

        Return theCustomerList

    End Function


End Class
