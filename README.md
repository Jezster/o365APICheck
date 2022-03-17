# O365 API checker

O365 query generator using the Microsoft worldwide.json dynamic resourse

Microsoft update this file to indicate modifications to URLs and IPs that relate to O365 locations

Sub-Filtering is based upon category:
  - Optimize
  - Allow
  - Default

## Operation
```
  usage: o365query.py [-h] [-a] [-o] [-d] [-6] service_area

  positional arguments:
    service_area    <Common | Exchange | SharePoint | Skype | All>

  optional arguments:
    -h, --help      show this help message and exit
    -a, --allow     Allow Category
    -o, --optimize  Optimize Category
    -d, --default   Default Category
    -6, --ipv6      Include IPv6 Addresses
```

## Example output

```
jerry@sunlight:~/Dev/O365$ python3 o365query.py -o Skype
Starting Collection...
File received
Generating content...

---------------------------------------------------------------------------
Created from worldwide.json with ID: 11
Service Category: Optimize
Service Area: Skype
Service Area Description: Skype for Business Online and Microsoft Teams
---------------------------------------------------------------------------
UDP Ports:      3478,3479,3480,3481
IPv4 Prefix:    13.107.64.0/18
IPv4 Prefix:    52.112.0.0/14
IPv4 Prefix:    52.120.0.0/14
---------------------------------------------------------------------------
Completed.
jerry@sunlight:~/Dev/O365$
```
