# Access the Paychex API, download the workers list, and merge the results with the Personnel List in the database.
# This is just a sample.  You must have a Client ID and Client Secret to a working API account with Paychex.
# You must also supply database connection strings, and alter to match your database requirements.  This is just a demonstration.
param(
    [switch]$test # test paychex api and database connectivity, but change no data
)
# sets security protocol for requests to tls 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#$paychex_uri = 'https://sandbox.api.paychex.com' # for sandbox play
$paychex_uri = 'https://api.paychex.com'
$client_id = '' # Client ID (API Key) assigned to this app for the Paychex API
$client_secret = '' # Client Secret (Key Secret) assigned to this app for the Paychex API
$comp_DisplayId = '11061545' # sandbox company
$sqlConnString = 'Data Source=Server\Instance;Initial Catalog=InitCatalog;Integrated Security=True'
$log_file = @{
    FilePath = 'Path To\Paychex.log'
    Append   = $true
    Encoding = 'UTF8'
}

"`r`n$(Get-Date) INFO Paychex Team Member Merge script starting!" | Out-File @log_file

# get authorization to the API
if (test-path variable:local:auth_result) { Remove-Variable auth_result }
$auth_result = Invoke-RestMethod -Method Post -Uri "$paychex_uri/auth/oauth/v2/token?grant_type=client_credentials&client_id=$client_id&client_secret=$client_secret"
if ($local:auth_result) {
    # create base authorization header for future requests
    $auth_header = @{'Authorization' = "$($auth_result.token_type) $($auth_result.access_token)" }

    # we shouldn't need to get a list of companies in the final app, because this app is only planned to have one company.
    #$companies_result = invoke-RestMethod -uri "$paychex_uri/companies" -headers $auth_header

    # find the companyId based on the displayId ($comp_DisplayId)
    if (test-path variable:local:companyId) { Remove-Variable companyId }
    $companyId = (Invoke-RestMethod -uri "$paychex_uri/companies?displayid=$comp_DisplayId" -headers $auth_header).content.companyId
    #$companyId = (Invoke-RestMethod -uri "$paychex_uri/companies?displayid=$comp_DisplayId" -Authentication Bearer -Token $auth_result.access_token).content.companyId  ## requires PS 6.0!
    if ($local:companyId) {

        # get the workers for the company found by display ID
        if (test-path variable:local:workers_result) { Remove-Variable workers_result }
        $workers_result = Invoke-RestMethod -uri "$paychex_uri/companies/$companyId/workers" -headers $auth_header
        if ($local:workers_result) {
            "$(Get-Date) INFO Paychex Workers Count=$($workers_result.content.Count)" | Out-File @log_file

            # create a table object to transfer the workers to the database team member merge procedure
            $TCEmpList = [Data.DataTable]::new()
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

            if (-not $test) {
                # send data table to stored procedure on the database to merge with Personnel List
                if (test-path variable:local:sqlResult) { Remove-Variable sqlResult }
                $sqlCmd = [Data.SqlClient.SqlCommand]::new('dbo.ImportTCEmployeeList', [Data.SqlClient.SqlConnection]::new($sqlConnString))
                $sqlCmd.CommandType = [Data.CommandType]::StoredProcedure
                $sqlCmd.Parameters.Add('@TCEmpList', [Data.SQLDBType]::Structured) | Out-Null
                $sqlCmd.Parameters['@TCEmpList'].Value = $TCEmpList
                $sqlCmd.Parameters['@TCEmpList'].TypeName = 'dbo.ut_tv_TCEmployeeList'
                ##$sqlCmd.Parameters['@TCEmpList'].Direction = [Data.SqlClient.SqlParameter]::Input  #  this is the default value
                $sqlCmd.Connection.Open()
                $sqlResult = $sqlCmd.ExecuteNonQuery()
                $sqlCmd.Connection.Close()

                "$(Get-Date) INFO SQL Merge Result=$sqlResult" | Out-File @log_file
            } else {
                # test transfer to database, validate user type definition, data transformations
                $sqlCmd = [Data.SqlClient.SqlCommand]::new('
                    SELECT tcel.TCEID tc_TCEID, pl.TCEID pl_TCEID, tcel.Inactive tc_Inactive, pl.Inactive pl_Inactive,
                    tcel.FirstName tc_FirstName, tcel.LastName tc_LastName, pl.[Team Member] FROM @TCEmpList tcel FULL OUTER JOIN 
                    dbo.[Personnel List] pl ON tcel.TCEID = pl.TCEID ORDER BY pl.[Team Member]',
                    [Data.SqlClient.SqlConnection]::new($sqlConnString))
                #$sqlCmd.CommandType = [Data.CommandType]::Text
                $sqlCmd.Parameters.Add('@TCEmpList', [Data.SQLDBType]::Structured) | Out-Null
                $sqlCmd.Parameters['@TCEmpList'].Value = $TCEmpList
                $sqlCmd.Parameters['@TCEmpList'].TypeName = "dbo.ut_tv_TCEmployeeList"
                $sqlAdapter = [Data.SqlClient.SqlDataAdapter]::new($sqlCmd)
                $sqlCmd.Connection.Open()
                $sqlAdapter.Fill(($sqlResultData = [Data.DataTable]::new()))
                $sqlCmd.Connection.Close()

                $sqlResultData.Rows # output the resulting data rows
                "$(Get-Date) INFO SQL Test Completed, $(if ($sqlResultData.Rows) {$sqlResultData.Rows.Count} else {'No'}) Merged Records Returned" | Out-File @log_file
            }
        } else {
            # no workers were returned, this is probably not right, $error should be recorded to a log
            "$(Get-Date) ERROR No result for Paychex workers request!  Error log follows:`r`n$error" | Out-File @log_file
        }
    } else {
        # could not find matching company ID, $error should be recorded to a log
        "$(Get-Date) ERROR No result for Paychex company ID ($comp_DisplayId) lookup!  Error log follows:`r`n$error" | Out-File @log_file
    }
} else {
    # could not get an authorization to the Paychex API, $error should be recorded to a log
    "$(Get-Date) ERROR No result for Paychex authorization request!  Error log follows:`r`n$error" | Out-File @log_file
}

"$(Get-Date) INFO Paychex Team Member Merge script terminating!" | Out-File @log_file
