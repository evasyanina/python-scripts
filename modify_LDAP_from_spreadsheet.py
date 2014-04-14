# This script modifies title, job type and manager of the employee in the LDAP directory
# using data provided in Google spreadsheet. In this particular script email address of the employee
# was used as filter for search and modification.

import gspread
import getpass
import ldap
import ldap.modlist

#pull data from the spreadsheet
password_google = getpass.getpass()
login_google = gspread.login("your_email_address", password_google)
google_doc = login_google.open("name_of_the_spreadsheet").worksheet("Sheet1")

#open a connection to the ldap
try:
        l = ldap.initialize("ldap://localhost.localdomain:port/")

#bind with a user to add objects
        username = "cn=Manager"
        password_ldap = getpass.getpass()
        l.simple_bind_s(username, password_ldap)

#loop through each entry in the spreadsheet to get data
	for i in range(1, 100): #set the range for entries in the spreadsheet
		email_string = []
		email_string.append('D') #column in the sreadsheet with email addresses
		email_string.append(str(i))
		email_joined = ''.join(email_string)	
		cell_email = google_doc.acell(email_joined).value
	
		title_string = []
		title_string.append('E') #column in the spreadsheet with employee title
		title_string.append(str(i))
		title_joined = ''.join(title_string)
		cell_title = google_doc.acell(title_joined).value

		type_string = []
		type_string.append('H') #column in the spreadsheet with job type
		type_string.append(str(i))
		type_joined = ''.join(type_string)
		cell_type = google_doc.acell(type_joined).value

		manager_string = []
                manager_string.append('G') #column in the spreadsheet with the employee manager
                manager_string.append(str(i))
                manager_joined = ''.join(manager_string)
                cell_manager = google_doc.acell(manager_joined).value

#employee search in the LDAP based on the email address to modify title
		baseDN = "ou=people, dc=example, dc=com"
		searchScope = ldap.SCOPE_SUBTREE
		retrieveAttributes = ['mail', 'title']
		searchFilter = "mail=%s" % cell_email
		
#perform search
		try:
			ldap_result_id = l.search(baseDN, searchScope, searchFilter, retrieveAttributes)
			while 1:
				result_type, result_data = l.result(ldap_result_id, 0)
				if (result_data == []):
					break
				else:
					if result_type == ldap.RES_SEARCH_ENTRY:
						result_name = result_data[0][0]
						result_title = result_data[0][1] 
		except ldap.LDAPError, e:
			print e
	
#the dn of our existing entry
		dn = "%s" % result_name
		
#some placeholders for old and new values
		old_title = {'title':result_title}
		new_title = {'title':cell_title}

#convert place-holders for modify-operation using modlist-module
		ldif_title = ldap.modlist.modifyModlist(old_title, new_title)

#do the actual modification
		l.modify_s(dn, ldif_title)

#print results
		#print "Finished title modification for %s" % result_name

#employee type search        
                baseDN = "ou=people, dc=example, dc=com"
                searchScope = ldap.SCOPE_SUBTREE
                retrieveAttributes = ['mail', 'employeeType']
                searchFilter = "mail=%s" % cell_email

#perform search
                try:
                        ldap_result_id = l.search(baseDN, searchScope, searchFilter, retrieveAttributes)
                        while 1:
                                result_type, result_data = l.result(ldap_result_id, 0)
                                if (result_data == []):
                                        break
                                else:
                                        if result_type == ldap.RES_SEARCH_ENTRY:
                                                result_name = result_data[0][0]
                                                result_employee = result_data[0][1]
                except ldap.LDAPError, e:
                        print e

#the dn of our existing entry
                dn = "%s" % result_name

#some placeholders for old and new values
                old_type = {'employeeType':result_employee}
                new_type = {'employeeType':cell_type}

#convert place-holders for modify-operation using modlist-module
                ldif_type = ldap.modlist.modifyModlist(old_type, new_type)

#do the actual modification
                l.modify_s(dn, ldif_type)

#print results
                #print "Finished employee modification for %s" % result_name

#manager search        
		baseDN = "ou=people, dc=example, dc=com"
		searchScope = ldap.SCOPE_SUBTREE
		retrieveAttributes = ['cn','mail']
		searchFilter = "mail=*%s*" % cell_manager

		try: 
			ldap_result_id = l.search(baseDN, searchScope, searchFilter, retrieveAttributes)
                        while 1:
                                result_type, result_data = l.result(ldap_result_id, 0)
                                if (result_data == []):
                                        break
                                else:
                                        if result_type == ldap.RES_SEARCH_ENTRY:
                                                result_manager_name = result_data[0][0]
		except ldap.LDAPError, e:
                        print e
		
                baseDN = "ou=people, dc=example, dc=com"
                searchScope = ldap.SCOPE_SUBTREE
                retrieveAttributes = ['mail', 'manager']
                searchFilter = "mail=*%s*" % cell_email

#perform search
                try:
                        ldap_result_id = l.search(baseDN, searchScope, searchFilter, retrieveAttributes)
                        while 1:
                                result_type, result_data = l.result(ldap_result_id, 0)
                                if (result_data == []):
                                        break
                                else:
                                        if result_type == ldap.RES_SEARCH_ENTRY:
                                                result_name = result_data[0][0]
                                                result_manager = result_data[0][1]
                        #print "Search result: %s" % result_name
                        #print "%s" %result_manager
                except ldap.LDAPError, e:
                        print e

#the dn of our existing entry
                dn = "%s" % result_name

#some placeholders for old and new values
                old_manager = {'manager':result_manager}
                new_manager = {'manager':result_manager_name}

#convert place-holders for modify-operation using modlist-module
                ldif_manager = ldap.modlist.modifyModlist(old_manager, new_manager)

#do the actual modification
                l.modify_s(dn, ldif_manager)

#print results
                #print "Finished manager modification: %s" % result_name
		print "Finished all modifications for %s" % result_name

#disconnect from the server and free resources
	l.unbind_s()

except ldap.LDAPError, e:
	print e
