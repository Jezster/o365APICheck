# O365 query generator using the Microsoft worldwide.json dynamic resourse
# Microsoft update this file to indicate modifications to URLs and IPs that relate to O365 locations
# Filtering for Skype (Teams)
#
# Sub-Filtering is based upon category:
#  - Optimize
#  - Allow
#  - Default

import requests
import re
import argparse

# Use Requests library to access the Microsoft Site to collect O365 worldwide.json 
def getMSReport(serviceArea):
    try:
        if serviceArea == "All":
            url = "https://endpoints.office.com/endpoints/Worldwide?ClientRequestId=b10c5ed1-bad1-445f-b386-b919946339a7"
        else:
            url = "https://endpoints.office.com/endpoints/worldwide?ServiceAreas={0}&clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7".format(serviceArea)
        payload = {}
        headers= {}
        try:
            response = requests.request("GET", url, headers=headers, data = payload)
            print("File received")
        except:
            print("ERROR: Connection failed to the Microsoft wordlwide.json URL: Check network connectivity")
            print("ERROR: Exiting.")
            exit()
        # Create json formated response
        json_response = response.json()
        # print(json_response)
        return json_response
    except:
        print('ERROR: Incorrect Service Area Entered - please reference the help with --help')
        exit()

# Create Proxy code to standard output 
def print_proxy(o365info, wants_optimize, wants_allow, wants_default, wants_ipv6):
    # Regex compile for IPv4 matching
    good_ip = re.compile("^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$")
    # Create a list that holds will hold the worldwide.json content
    complete_list = []
    # Assign worldwide.json content to the list for processing - each record is defined as a list with a record holding key information
    complete_list = o365info
    for record in complete_list:   
        # Identify if a Category of traffic is required and then generate Proxy Script
        if (record["category"] == "Optimize" and wants_optimize == True) or \
           (record["category"] == "Allow" and wants_allow == True) or \
           (record["category"] == "Default" and wants_default == True):
            # Identify that a list of Proxy lines are being generated - to ensure that a close bracket is printed
            all_urls = []
            all_ips = []
            # Generate header for the Proxy file
            print ("")
            print ("---------------------------------------------------------------------------")
            print ("Created from worldwide.json with ID: {0}".format(str(record["id"])))
            print ("Service Category: {0} ".format(record["category"]))
            print ("Service Area: {0}".format(record["serviceArea"]))
            print ("Service Area Description: {0}".format(record["serviceAreaDisplayName"]))
            print ("---------------------------------------------------------------------------")
            tcpports = record.get("tcpPorts")
            if tcpports:
                print("TCP Ports:\t{0} ".format(record["tcpPorts"]))
            udpports = record.get("udpPorts")
            if udpports:
                print("UDP Ports:\t{0}".format(record["udpPorts"]))
            notes = record.get("notes")
            if notes:
                print ("Available Notes:\t{0}".format(record["notes"]))
            urls = record.get("urls")
            # URL address line loop
            if urls:
                for url in urls:
                    all_urls.append("URL: \t\t{0}".format(url))
            # IP address line loop
            ips = record.get("ips")
            if ips:
                for ip in ips:
                    ipaddr =  ip.split("/",1)
                    test = good_ip.match(ipaddr[0])
                    if test:
                        all_urls.append("IPv4 Prefix:\t{0}/{1}".format(ipaddr[0], ipaddr[1]))
                    elif wants_ipv6 == True:
                        all_urls.append("IPv6 Prefix:\t{0}/{1}".format(ipaddr[0], ipaddr[1]))
            # Build automatic lines to allow the close of the If( statement in Proxy
            total_proxy_list = all_ips + all_urls
            counter = 0
            list_length = len(total_proxy_list)
            for item in total_proxy_list:
                counter+= 1
                if counter == list_length:
                    print ("{0}".format(item))
                else:
                    print ("{0}".format(item))
            
def parse_cli():
    parser = argparse.ArgumentParser()
    parser.add_argument('service_area', help='<Common | Exchange | SharePoint | Skype | All> ', type=str)
    parser.add_argument('-a', '--allow', help='Allow Category', action="store_true")
    parser.add_argument('-o', '--optimize', help='Optimize Category', action="store_true")
    parser.add_argument('-d', '--default', help='Default Category', action="store_true")
    parser.add_argument('-6', '--ipv6', help="Include IPv6 Addresses", action="store_true")
    args = parser.parse_args()
    return args

# Main calls
def main():
    args = parse_cli()
    print ("Starting Collection...")
    o365info = getMSReport(args.service_area.capitalize())
    # Create an empty list that will contain the output from the Microsft worldwide.json response
    # The microsoft JSON response contains a list with each object referencing an ID
    print ("Generating content...")
    print_proxy(o365info, args.optimize, args.allow, args.default, args.ipv6)
    print ("---------------------------------------------------------------------------")
    print ("Completed.")


if __name__ == "__main__":
    # execute only if run as a script
    main()
