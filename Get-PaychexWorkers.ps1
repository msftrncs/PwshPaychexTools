# Access the Paychex API, download the workers list, and merge the results with the Personnel List in the database.
# This is just a sample.  You must have a Client ID and Client Secret to a working API account with Paychex.
# You must also supply database connection strings, and alter to match your database requirements.  This is just a demonstration.

# sets security protocol for requests to tls 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#$paychex_uri = 'https://sandbox.api.paychex.com' # for sandbox play
$paychex_uri = 'https://api.paychex.com'
$client_id = '' # Client ID from the Paychex API account
$client_secret = '' # Client ID from the Paychex API account
$comp_DisplayId = '11061545' # sandbox company
$sqlConnString = 'Data Source=Server\Instance;Initial Catalog=InitCatalog;Integrated Security=True'
$log_file = 'Path To\Paychex.log'

"`r`n$(Get-Date)-Paychex Team Member Merge script starting!" | Out-File $log_file -Append

# get authorization to the API
if (test-path variable:auth_result) {Remove-Variable auth_result}
$auth_result = Invoke-RestMethod -Method Post -Uri "$paychex_uri/auth/oauth/v2/token?grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret"
if ($auth_result) {
    # create base authorization header for future requests
    $auth_header = @{'Authorization' = "$($auth_result.token_type) $($auth_result.access_token)"}

    # we shouldn't need to get a list of companies in the final app, because this app is only planned to have one company.
    #$companies_result = invoke-RestMethod -uri "${paychex_uri}/companies" -headers $auth_header

    # find the companyId based on the displayId ($comp_DisplayId)
    if (test-path variable:companyId) {Remove-Variable companyId}
    $companyId = (Invoke-RestMethod -uri "$paychex_uri/companies?displayid=$comp_DisplayId" -headers $auth_header).content.companyId
    #$companyId = (Invoke-RestMethod -uri "${paychex_uri}/companies?displayid=${comp_DisplayId}" -Authentication Bearer -Token ${auth_result}.access_token).content.companyId  ## requires PS 6.0!
    if ($companyId) {

        # get the workers for the company found by display ID
        if (test-path variable:workers_result) {Remove-Variable workers_result}
        $workers_result = Invoke-RestMethod -uri "$paychex_uri/companies/$companyId/workers" -headers $auth_header
        if ($workers_result) {
            "$(Get-Date)-Paychex Workers Count=$($workers_result.content.Count)" | Out-File $log_file -Append

            # create a table object to transfer the workers to the database team member merge procedure
            $TCEmpList = New-Object 'Data.DataTable'
            $TCEmpList.PrimaryKey = $TCEmpList.Columns.Add('TCEID', [Int32])
            $TCEmpList.Columns.Add('FirstName', [String]) | Out-Null
            $TCEmpList.Columns.Add('LastName', [String]) | Out-Null
            $TCEmpList.Columns.Add('Inactive', [Boolean]) | Out-Null

            # translate each worker record from web result to data table
            ForEach ($worker in $workers_result.Content) {
                $row = $TCEmpList.NewRow()
                $row['TCEID'] = $worker.employeeId
                $row['FirstName'] = $worker.name.givenName
                $row['LastName'] = $worker.name.familyName
                $row['Inactive'] = $worker.currentStatus.statusType -ne 'ACTIVE'
                $TCEmpList.Rows.Add($row)
            }

            # send data table to stored procedure on the database to merge with Personnel List
            if (test-path variable:sqlResult) {Remove-Variable sqlResult}
            $sqlcmd = New-Object 'Data.SqlClient.SqlCommand' -ArgumentList 'dbo.ImportTCEmployeeList'
            $sqlcmd.CommandType = [Data.CommandType]::StoredProcedure
            $sqlcmd.Parameters.Add('@TCEmpList', [Data.SQLDBType]::Structured) | Out-Null
            $sqlcmd.Parameters['@TCEmpList'].Value = $TCEmpList
            $sqlcmd.Parameters['@TCEmpList'].TypeName = 'dbo.ut_tv_TCEmployeeList'
            ##$sqlcmd.Parameters['@TCEmpList'].Direction = 'Input'  #  this is the default value
            $sqlcmd.Connection = New-Object 'Data.SqlClient.SqlConnection' -ArgumentList $sqlConnString
            $sqlcmd.Connection.Open()
            $sqlResult = $sqlcmd.ExecuteNonQuery()
            $sqlcmd.Connection.Close()

            "$(Get-Date)-SQL Merge Result=$sqlResult" | Out-File $log_file -Append

            <# ## test transfer to database, validate user type definition, data transformations
            $dt = New-Object 'Data.DataTable'
            $sqlcmd = New-Object 'System.Data.SqlClient.SqlCommand' -ArgumentList 'SELECT tcel.TCEID, pl.TCEID, tcel.Inactive, pl.Inactive, tcel.FirstName, tcel.LastName, pl.[Team Member] FROM @TCEmpList tcel LEFT JOIN dbo.[Personnel List] pl ON tcel.TCEID = pl.TCEID ORDER BY pl.[Team Member]'
            #$sqlcmd.CommandType = [Data.CommandType]::Text
            $sqlcmd.Parameters.Add('@TCEmpList',[Data.SQLDBType]::Structured) | Out-Null
            $sqlcmd.Parameters['@TCEmpList'].Value = $TCEmpList
            $sqlcmd.Parameters['@TCEmpList'].TypeName = "dbo.ut_tv_TCEmployeeList"
            $sqlcmd.Connection = New-Object 'Data.SqlClient.SqlConnection' -ArgumentList $sqlConnString
            $sqlcmd.Connection.Open()
            $da = New-Object 'System.Data.SqlClient.SqlDataAdapter' -ArgumentList $sqlcmd
            $da.Fill($dt)
            $sqlcmd.Connection.Close()
            #>
        }
        else {
            # no workers were returned, this is probably not right, $error should be recorded to a log
            "$(Get-Date)-No result for Paychex workers request!  Error log follows:`r`n$error" | Out-File $log_file -Append
        }
    }
    else {
        # could not find matching company ID, $error should be recorded to a log
        "$(Get-Date)-No result for Paychex company ID ($comp_DisplayId) lookup!  Error log follows:`r`n$error" | Out-File $log_file -Append
    }
}
else {
    # could not get an authorization to the Paychex API, $error should be recorded to a log
    "$(Get-Date)-No result for Paychex authorization request!  Error log follows:`r`n$error" | Out-File $log_file -Append
}

"$(Get-Date)-Paychex Team Member Merge script terminating!" | Out-File $log_file -Append
