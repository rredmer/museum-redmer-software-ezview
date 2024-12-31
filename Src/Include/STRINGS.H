***************************************************
**Sams Mailer Program Header
***************************************************
**STRINGS.H
**Purpose: Precompile definitions.
***************************************************
**Author: Barrie Sosinsky
**Killer Apps, Inc., Copyright 1995
**44 Mechanic St., Newton MA 02164
**(617)244-8100 tel, (617)244-6344 fax
**basman@killerapps.com; CIS: 72020,2311
**Date Created: March 28, 1995
**Date Last Modified: April 23, 1995
**Last Modified By: Barrie Sosinsky
**Modification Notes: Used for localization. These
**strings are mostly copyrighted by Microsoft, Inc.
***************************************************

*-- Messagebox Titles
#DEFINE ERRORTITLE_LOC			"An error has occurred"
#DEFINE MAILMAN_LOC				"MailMan"

*-- Messagebox Messages
#DEFINE FILENOTEXIST_LOC		"File does not exist: "
#DEFINE BADPASSWORD_LOC			"Password is invalid."
#DEFINE BADNAME_LOC				"Name not found."
#DEFINE BADUPDATE_LOC			"Could not update - reverting to original."
#DEFINE SAVECHANGES_LOC			"Do you want to save your changes first?"
#DEFINE ASKDELETE_LOC			"Are you sure you wish to delete this information?"
#DEFINE TAGNOTFOUND_LOC			"Index tag not found."
#DEFINE REPORTNOTFOUND_LOC		"Report file not found."
#DEFINE PRINTERNOTREADY_LOC		"Printer not ready."
#DEFINE NORECSMATCHED_LOC		"No records matched criteria."
#DEFINE DELETEREC_LOC			"Are you sure you want to delete this record?"
#DEFINE DELETEWARN_LOC			"Delete Warning"
#DEFINE AVAILABLECREDIT_LOC		"Available Credit"
#DEFINE NOTYET_LOC				"Under Construction"
#DEFINE CUSTIDEXISTS_LOC		"Customer ID already exists. Please re-enter."
#DEFINE NOLASTORDER_LOC			"Customer has no prior order."
#DEFINE CUSTFIRSTORDER_LOC		"Customer's last order is current order."
#DEFINE TODAYORLATER_LOC		"Date must be today or later."
#DEFINE DATERANGEERROR_LOC		"'To' date cannot be less than 'From' date."
#DEFINE ADDCUSTOMER_LOC			"Do you want to add this customer to the Customer master file?"
#DEFINE NOTHINGTOPRINT_LOC		"Nothing to print."
#DEFINE PASSWORDEMPTY_LOC		"New password cannot be empty."
#DEFINE PSWDNOTCNFRM_LOC		"Cannot confirm new password. Please try again."
#DEFINE NOPSWDENTERED_LOC		"You have not yet entered the old password. Do you want to continue?"
#DEFINE FILESAVEDAS_LOC			"File saved as "
#DEFINE DELSHIPPER_LOC			"Shippers exists on orders. Cannot delete."
#DEFINE ORDHASITEMS_LOC			"An order must have at least one line item."
#DEFINE CUSTOVERMAX_LOC			"Customer is over their maximimun order amount by "
#DEFINE CUSTUNDERMIN_LOC		"Customer order total must be at least "
#DEFINE SAVEANYWAY_LOC			"Save anyway?"
#DEFINE VALIDATING_LOC			"Validating ..."
#DEFINE VIEWCODEPRINT_LOC		"This report may be lenghty. Do you want to continue?"
#DEFINE VIEWCSDTYPRINT_LOC		"This report may be lenghty. Do you want to continue?"
#DEFINE PRINTING_LOC				"Printing ..."
#DEFINE METHOD_LOC					"Method: "
#DEFINE LINENUM_LOC					"Line: "
#DEFINE CUSTNOORD_LOC				"Customer has no orders."
#DEFINE SELCUSTFIRST_LOC		"Must select a customer first."
#DEFINE CANNOTQUIT_LOC			"Cannot quit Visual FoxPro within Sams Mailer."
#DEFINE ADDNEWREC_LOC				"That was the last record. Do you want to add a new one?"
#DEFINE ENTERADDMODE_LOC		"There are no records on file. You will be placed in 'Add' mode."
#DEFINE NOEMPLOYEES_LOC			"There are no employees on file."
#DEFINE INSEMPLOYEE_LOC			"All employees must be assigned to a group."
#DEFINE INSPRODUCT_LOC			"All products must be assigned a supplier and a category."
#DEFINE INSORDER_LOC				"All orders must have a customer and a shipper. "
#DEFINE TABLERULEFAIL_LOC		"Table rule failed!"

*-- Trigger error messages
#DEFINE INSERTTRIGFAIL_LOC		"Insert trigger failed!"
#DEFINE UPDATETRIGFAIL_LOC		"Update trigger failed!"
#DEFINE DELETETRIGFAIL_LOC		"Delete trigger failed!"
#DEFINE DELCATEGORY_LOC			"Products belong to this category. Cannot delete!"
#DEFINE DELCUSTOMER_LOC			"Customer has orders. Cannot delete!"
#DEFINE DELEMPLOYEE_LOC			"Employee exists on orders. Cannot delete!"
#DEFINE DELPRODUCT_LOC			"Product exists on order line items. Cannot delete!"

*-- Other strings
#DEFINE ADDPICTURE_LOC			"Add Picture"
#DEFINE CHANGEPICTURE_LOC		"Change Picture"
#DEFINE SELECTBUTTON_LOC		"Select"
#DEFINE VERSION_LOC				"1.0"
#DEFINE COPYRIGHT_LOC			"Copyright 1996 Killer Apps, Inc."
#DEFINE RIGHTSRSRVD_LOC			"All rights reserved"
#DEFINE ADDITEM_LOC				"Add Item"
#DEFINE REMOVEITEM_LOC			"Remove Item"
#DEFINE INFSAVED_LOC					"Information saved."
#DEFINE REINDEXING_LOC				"Reindexing: "
#DEFINE DONE_LOC							"Done"
#DEFINE LOWERNOTFOUND_LOC			" not found."
#DEFINE ABOUT_LOC							"About "
#DEFINE VERSIONLABEL_LOC			"Version "
#DEFINE LOWERFOR_LOC					" for "
#DEFINE UPDATEORDER_LOC				"Updating Order"