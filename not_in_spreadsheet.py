# This script lists (and enters into Google spreadsheet) email addresses of all employees
# from the LDAP directory that are not entered in Google spreadsheet. In this particular script
# email address of the employee was used as filter for search and modification.

import gspread
import getpass
import ldap
import ldap.modlist

#pull data from the spreadsheet
password_google = getpass.getpass()
login_google = gspread.login("email_address@gmail.com", password_google)
google_doc = login_google.open("name_of_the_spreadsheet").worksheet("Sheet1")

#open a connection to the ldap
try:
        l = ldap.initialize("ldap://localhost.localdomain:port/")

#bind with a user to add objects
        username = "cn=Manager"
        password_ldap = getpass.getpass()
	l.simple_bind_s(username, password_ldap)

#search basics
	baseDN = "ou=people, dc=example, dc=com"
        searchScope = ldap.SCOPE_SUBTREE
        retrieveAttributes = ['mail']
        searchFilter = "mail=*"

#perform search
	try:
        	ldap_result_id = l.search(baseDN, searchScope, searchFilter, retrieveAttributes)
		result_set = []         
		j = 2 # first row with data in the spreadsheet             
 		while 1:
                	result_type, result_data = l.result(ldap_result_id, 0)
                        if (result_data == []):
                        	break
                        else:
                                if result_type == ldap.RES_SEARCH_ENTRY:
    					result_set.append(result_data)
		for i in range(len(result_set)):
			for entry in result_set[i]:			
				try:
					email = entry[1]['mail'][0]
					values = google_doc.col_values(4)
					if email not in values:
						#print email
						email_string = []
                				email_string.append('J') # print results in column "J" of the spreadsheet
						email_string.append(str(j))
						email_update = ''.join(email_string)
						google_doc.update_acell(email_update, email)
						j += 1		
				except:
					pass
	except ldap.LDAPError, e:
                print e

#disconnect from the server and free resources
        l.unbind_s()

except ldap.LDAPError, e:
        print e
