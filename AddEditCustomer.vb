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
Imports Newtonsoft.Json
Imports System.IO

Partial Public Class CustomersController
    Inherits ApiController


    ' POST api/<app>/<controller>/<id>
    ''' <summary>
    ''' Adds a new customer or updates an existing customer, including shipping address
    ''' </summary>
    ''' <param name="app">Name of application</param>
    ''' <param name="value">NameValue collection for forms</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <ApiAuthorizeAttribute>
    Public Function AddEditCustomer(ByVal app As String, <FromBody()> ByVal obj As String) As HttpResponseMessage

        Dim lstErrors As New List(Of ErrorsDetails)

        HttpContext.Current.Request.InputStream.Seek(0, SeekOrigin.Begin)
        Dim jsonObj As String = New StreamReader(HttpContext.Current.Request.InputStream).ReadToEnd()
        obj = jsonObj

        'Check that JSON body exists. If not, exit function and return error.
        If obj Is Nothing Then
            Dim Result As New ErrorsDetails
            Result.Name = "JSON Body"
            Result.Value = "JSON Body is required."
            lstErrors.Add(Result)

            Dim Err As New Errors
            Err.Message = lstErrors

            Dim response As HttpResponseMessage = New HttpResponseMessage(HttpStatusCode.BadRequest)
            response.Content = New StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(Err))
            Return response

        End If
        Dim value As New CustomerUpdateObject
        value = JsonConvert.DeserializeObject(Of CustomerUpdateObject)(obj.ToString())
        Dim values As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(obj.ToString())

        Dim fields = New List(Of CustomField)()

        For Each Val As KeyValuePair(Of String, Object) In values
            If Not PropertyUtils.Exists(Val.Key, value) And Not Val.Value Is Nothing Then
                Dim field = New CustomField()
                'field.CustomerId = customer.CustomerId, 'MUST LINK TO CUSTOMER HERE
                field.FieldName = Val.Key
                field.FieldValue = Val.Value.ToString()
                fields.Add(field)
            End If
        Next
        value.CustomFieldList = fields



        ''required



        If value.CustomerCountry.Trim = String.Empty Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerCountry"
            Result.Value = "Country is required."
            lstErrors.Add(Result)
        End If


        If value.CustomerFullName.Trim = String.Empty Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerFullName"
            Result.Value = "FullName is required."
            lstErrors.Add(Result)
        ElseIf Not Regex.IsMatch(value.CustomerFullName.Trim, AppSettings("RegexFullName")) Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerFullName"
            Result.Value = "Invalid characters in FullName."
            lstErrors.Add(Result)
        End If


        If value.CustomerPhone.Trim = String.Empty Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerPhone"
            Result.Value = "Phone is required."
            lstErrors.Add(Result)
        Else

            If value.CustomerCountry = "United States" Then
                'US Phone

                'Parse Phone Number
                Dim thePhoneUSA As String = value.CustomerPhone.Trim

                Dim sb = New StringBuilder()

                For Each c As Char In thePhoneUSA
                    If Char.IsNumber(c) Then
                        sb.Append(c)
                    End If
                Next

                thePhoneUSA = sb.ToString()

                'Remove Leading 1
                If Left(thePhoneUSA, 1) = "1" Then
                    thePhoneUSA = thePhoneUSA.Substring(1)
                End If

                'Get the first 10 characters
                If IsNumeric(thePhoneUSA) Then
                    If thePhoneUSA.Length > 9 Then
                        thePhoneUSA = Left(thePhoneUSA, 10)
                    End If
                End If

                If Not Regex.IsMatch(thePhoneUSA, AppSettings("RegexPhoneUSA")) Then
                    Dim Result As New ErrorsDetails
                    Result.Name = "CustomerPhone"
                    Result.Value = "Invalid characters in Phone."
                    lstErrors.Add(Result)
                End If

                value.CustomerPhone = thePhoneUSA

            Else
                'Intl Phone
                If Not Regex.IsMatch(value.CustomerPhone.Trim, AppSettings("RegexPhoneIntl")) Then
                    Dim Result As New ErrorsDetails
                    Result.Name = "CustomerPhoneIntl"
                    Result.Value = "Invalid characters in Intl Phone."
                    lstErrors.Add(Result)
                End If
            End If
        End If



        If value.CustomerEmail.Trim = String.Empty Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerEmail"
            Result.Value = "Email is required."
            lstErrors.Add(Result)
        ElseIf Not Regex.IsMatch(value.CustomerEmail.Trim, AppSettings("RegexEmail")) Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerEmail"
            Result.Value = "Invalid characters in Email."
            lstErrors.Add(Result)
        End If


        If value.CustomerAddress1.Trim = String.Empty Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerAddress1"
            Result.Value = "Address Line 1 is required."
            lstErrors.Add(Result)
        ElseIf Not Regex.IsMatch(value.CustomerAddress1.Trim, AppSettings("RegexStreetAddress")) Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerAddress1"
            Result.Value = "Invalid characters in Address Line 1."
            lstErrors.Add(Result)
        End If




        If value.CustomerCity.Trim = String.Empty Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerCity"
            Result.Value = "City is required."
            lstErrors.Add(Result)
        ElseIf Not Regex.IsMatch(value.CustomerCity.Trim, AppSettings("RegexCity")) Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerCity"
            Result.Value = "Invalid characters in City."
            lstErrors.Add(Result)
        End If



        'Validate State/Province for certain countries
        Select Case value.CustomerCountry.ToLower
            Case "united states", "mexico", "canada"

                Dim Result As New ErrorsDetails
                If value.CustomerState.Trim = String.Empty Then
                    Result.Name = "CustomerState"
                    Result.Value = "State is required."
                    lstErrors.Add(Result)
                ElseIf Not Regex.IsMatch(value.CustomerState.Trim, AppSettings("RegexCity")) Then
                    Result.Name = "CustomerState"
                    Result.Value = "Invalid characters in State."
                    lstErrors.Add(Result)
                End If

                'State/Provine
                value.CustomerState = value.CustomerState.Trim.ToUpper

        End Select


        If value.CustomerZip.Trim = String.Empty Then
            If ("us,usa,united states").Contains(value.CustomerCountry.Trim.ToLower()) Then
                Dim Result As New ErrorsDetails
                Result.Name = "CustomerZip"
                Result.Value = "Zip Code is required."
                lstErrors.Add(Result)
            Else
                Dim Result As New ErrorsDetails
                Result.Name = "CustomerZip"
                Result.Value = "Postal Code is required."
                lstErrors.Add(Result)
            End If
        ElseIf ("us,usa,united states").Contains(value.CustomerCountry.Trim.ToLower()) And Not Regex.IsMatch(value.CustomerZip.Trim, AppSettings("RegexZipUSA")) Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerZip"
            Result.Value = "Invalid Zip Code format."
            lstErrors.Add(Result)
        ElseIf Not ("us,usa,united states").Contains(value.CustomerCountry.Trim.ToLower()) And Not Regex.IsMatch(value.CustomerZip.Trim, AppSettings("RegexZipIntl")) Then
            Dim Result As New ErrorsDetails
            Result.Name = "CustomerZip"
            Result.Value = "Invalid Postal Code format."
            lstErrors.Add(Result)
        End If


        'FullName
        value.CustomerFullName = Helpers.FixCase(value.CustomerFullName).Trim



        'Organization
        If Not value.CustomerOrganization.Trim = String.Empty Then
            If Regex.IsMatch(value.CustomerOrganization.Trim, AppSettings("RegexOrgBrandName")) Then
                value.CustomerOrganization = Helpers.FixCase(value.CustomerOrganization).Trim
            Else
                Dim Result As New ErrorsDetails
                Result.Name = "CustomerOrganization"
                Result.Value = "Invalid characters in Organization."
                lstErrors.Add(Result)
            End If
        End If

        'Recipient
        If Not value.CustomerRecipient = Nothing And Not value.CustomerRecipient.Trim = String.Empty Then
            If Regex.IsMatch(value.CustomerRecipient.Trim, AppSettings("RegexOrgBrandName")) Then
                value.CustomerRecipient = Helpers.FixCase(value.CustomerRecipient).Trim
            Else
                Dim Result As New ErrorsDetails
                Result.Name = "CustomerRecipient"
                Result.Value = "Invalid characters in Recipient."
                lstErrors.Add(Result)
            End If
        Else
            value.CustomerRecipient = value.CustomerFullName
        End If


        'Email
        value.CustomerEmail = value.CustomerEmail.ToLower

        'Address1
        value.CustomerAddress1 = Helpers.FixCase(value.CustomerAddress1.Trim)

        'Address2
        If Not value.CustomerAddress2.Trim = String.Empty Then
            If Regex.IsMatch(value.CustomerAddress2.Trim, AppSettings("RegexStreetAddress")) Then
                value.CustomerAddress2 = Helpers.FixCase(value.CustomerAddress2.Trim)
            Else
                Dim Result As New ErrorsDetails
                Result.Name = "CustomerAddress2"
                Result.Value = "Invalid characters in Address Line 2."
                lstErrors.Add(Result)
            End If
        End If

        'City
        value.CustomerCity = Helpers.FixCase(value.CustomerCity.Trim)



        'Intl Zip
        If Not value.CustomerCountry = "United States" Then
            value.CustomerZip = value.CustomerZip.Trim.ToUpper
        End If



        '' Return IF any requuired or mising Invalid fields
        If lstErrors.Count > 0 Then
            Dim Err As New Errors
            Err.Message = lstErrors
            Dim response As HttpResponseMessage = New HttpResponseMessage(HttpStatusCode.BadRequest)
            response.Content = New StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(Err))
            Return response
        End If


        'Address Check IF CustomerAddressCheckOff is NOT True

        If value.CustomerCountry = "United States" And Not value.CustomerAddressCheckOff Then
            'Verify Address USPS
            Try
                'Get Validated Address
                Dim theAddress As New Address
                theAddress = Helpers.ValidateAddressUSPS(value.CustomerAddress1, value.CustomerAddress2, value.CustomerCity, value.CustomerState, "18901")

                'Update Customer Address
                value.CustomerAddress1 = Helpers.FixCase(theAddress.Address2)
                value.CustomerAddress2 = Helpers.FixCase(theAddress.Address1)
                value.CustomerCity = Helpers.FixCase(theAddress.City)

                'Add Plus 4 if not blank
                If Not theAddress.ZipPlus4.Trim = String.Empty Then
                    value.CustomerZip = theAddress.Zip & "-" & theAddress.ZipPlus4.Trim
                Else
                    value.CustomerZip = theAddress.Zip
                End If

            Catch ex As Exception

                If ex.Message.ToLower.Contains("address not found") Or ex.Message.ToLower.Contains("invalid address") Then
                    Dim Result As New ErrorsDetails
                    Result.Name = "AddressVerification"
                    Result.Value = "Shipping Address failed verification. If address is known to be valid, try alternate abbreviations and punctuation."
                    lstErrors.Add(Result)
                ElseIf ex.Message.ToLower.Contains("multiple addresses") Then
                    Dim Result As New ErrorsDetails
                    Result.Name = "AddressVerification"
                    Result.Value = "Shipping Address failed verification. More information such as Apt or Suite # may be needed."
                    lstErrors.Add(Result)

                ElseIf ex.Message.ToLower.Contains("invalid city") Then
                    Dim Result As New ErrorsDetails
                    Result.Name = "AddressVerification"
                    Result.Value = "The City is invalid for this State. If city is known to be valid, try alternate abbreviation and punctuation."
                    lstErrors.Add(Result)
                Else
                    ''Change By Chirag
                    Dim Result As New ErrorsDetails
                    Result.Name = "Other"
                    Result.Value = ex.Message.ToString
                    lstErrors.Add(Result)
                End If

            End Try

        End If


        If lstErrors.Count > 0 Then
            Dim Err As New Errors
            Err.Message = lstErrors

            Dim response As HttpResponseMessage = New HttpResponseMessage(HttpStatusCode.BadRequest)
            response.Content = New StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(Err))
            Return response
        End If



        'Get or Add Customer *********************************************************************************************************************

        theCustomerManager = New CustomerManager(UserManager.GetUserId, OrganizationManager.GetOrganizationId)
        theProvider = Membership.Providers("AspNetSqlMembershipProvider3")

        Dim theCustomerGuid As New Guid
        theCustomerGuid = Guid.Empty

        Try

            'Get Customer by Email
            Dim theCustomerFindByEmail As New Customer
            theCustomerFindByEmail = theCustomerManager.GetCustomerByEmail(value.CustomerEmail)

            If Not theCustomerFindByEmail Is Nothing Then
                theCustomerGuid = theCustomerFindByEmail.CustomerUserId
            End If


            'If customer still not found, Add NEW Customer 
            If theCustomerGuid = Guid.Empty Then

                theCustomerGuid = AddNewCustomer(value)

                'Add the Profile Info. If the ID is still empty, FAIL
                If Not theCustomerGuid = Guid.Empty Then
                    InsertCustomerProfileInfo(theCustomerGuid, value)
                    'theCustomerManager.InsertIdForApp(theCustomerGuid, app, value.CustomerId)
                Else
                    Return New HttpResponseMessage(HttpStatusCode.BadRequest)
                End If

            End If


            'Add the Customer to the Brand
            If Not AddCustomerToBrand(theCustomerGuid, value.CustomerEmail) Then
                Return New HttpResponseMessage(HttpStatusCode.BadRequest)
            End If

            'Update the Address 
            If Not UpdateCustomerAddress(theCustomerGuid, value) Then
                Return New HttpResponseMessage(HttpStatusCode.BadRequest)
            End If

            'Update the Contact 
            If Not UpdateCustomerContact(theCustomerGuid, value) Then
                Return New HttpResponseMessage(HttpStatusCode.BadRequest)
            End If



        Catch ex As Exception

            'Send Email on Database Error
            Email.SendErrorMail("Error in CustomerUpdateController", ex)

            Return New HttpResponseMessage(HttpStatusCode.InternalServerError)

        End Try

        Return New HttpResponseMessage(HttpStatusCode.OK)

    End Function



End Class
