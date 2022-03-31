I use Microsoft PowerApps to feed PowerShell runbooks in Automation Acounts in Azure.
With PowerAutomate the input will be parsed to JSON and the variables will be used to feed the PowerShell runbooks.

I use an automation account to perform the actions. In some runbooks I use Microsoft Graph. To run those actions I created a service principal account.
Give it the necessary API permissions to perform the tasks.