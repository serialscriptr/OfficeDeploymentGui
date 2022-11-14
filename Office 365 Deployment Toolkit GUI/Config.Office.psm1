<#
.DESCRIPTION: Very small module for getting some information from the config.office.com website
.AUTHOR: Ryan (SerialScripter)
#>

function Get-OfficeReleaseChannel
{
	Invoke-RestMethod -UseBasicParsing -Uri "https://clients.config.office.net/releases/v1.0/OfficeReleases" `
					  -Headers @{
		"authority"				    = "clients.config.office.net"
		"method"				    = "GET"
		"path"					    = "/releases/v1.0/OfficeReleases"
		"scheme"				    = "https"
		"accept"				    = "application/json"
		"accept-encoding"		    = "gzip, deflate, br"
		"accept-language"		    = "en-US,en;q=0.8"
		"origin"				    = "https://config.office.com"
		"referer"				    = "https://config.office.com/"
		"sec-fetch-dest"		    = "empty"
		"sec-fetch-mode"		    = "cors"
		"sec-fetch-site"		    = "cross-site"
		"sec-gpc"				    = "1"
		"x-api-name"			    = "https://clients.config.office.net/releases/v1.0/- api name not register"
		"x-manageoffice-client-sid" = "c2e5083a-541c-41ed-951f-5caedbbb51ae"
		"x-start-time"			    = "1667256169739"
	}
}

function Get-OfficeReleaseChannelForProduct
{
	param
	(
		[parameter(Mandatory = $true, HelpMessage = "The productId of the product you want to get release channels for")]
		[string]$ProductId
	)
	
	Invoke-RestMethod -UseBasicParsing -Uri "https://config.office.com/api/filelist/channelsForProductIds?productId=$ProductId" `
					  -Headers @{
		"authority"				    = "config.office.com"
		"method"				    = "GET"
		"path"					    = "/api/filelist/channelsForProductIds?productId=O365BusinessRetail"
		"scheme"				    = "https"
		"accept"				    = "application/json"
		"accept-encoding"		    = "gzip, deflate, br"
		"accept-language"		    = "en-US,en;q=0.8"
		"referer"				    = "https://config.office.com/deploymentsettings"
		"sec-fetch-dest"		    = "empty"
		"sec-fetch-mode"		    = "cors"
		"sec-fetch-site"		    = "same-origin"
		"sec-gpc"				    = "1"
		"x-api-name"			    = "fileListChannelForProductIds"
		"x-manageoffice-client-sid" = "c2e5083a-541c-41ed-951f-5caedbbb51ae"
		"x-start-time"			    = "1667256072919"
	}
}

function Get-OfficeProductLanguagePack
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$ProductId
	)
	
	Invoke-RestMethod -UseBasicParsing -Uri "https://config.office.com/api/filelist/languagesForProductIds?productId=$ProductId&languageType=ProofingOnly" `
					  -Headers @{
		"authority"				    = "config.office.com"
		"method"				    = "GET"
		"path"					    = "/api/filelist/languagesForProductIds?productId=O365ProPlusRetail&languageType=ProofingOnly"
		"scheme"				    = "https"
		"accept"				    = "application/json"
		"accept-encoding"		    = "gzip, deflate, br"
		"accept-language"		    = "en-US,en;q=0.8"
		"referer"				    = "https://config.office.com/deploymentsettings"
		"sec-fetch-dest"		    = "empty"
		"sec-fetch-mode"		    = "cors"
		"sec-fetch-site"		    = "same-origin"
		"sec-gpc"				    = "1"
		"x-api-name"			    = "api/filelist/languagesforproductids- api name not register"
		"x-manageoffice-client-sid" = "c2e5083a-541c-41ed-951f-5caedbbb51ae"
		"x-start-time"			    = "1667256780934"
	}
}