#######################################################
# This has to be run with the cert in the same folder #
#######################################################

# This Script does the following:

    # reads the .pfx input certificate
	# asks for the corresponding password
	# outputs the same certificate, but as a .der file

# Made by David S. 06.07.22
# Version 1.0



################################
# Code, nothing to change here #
################################
Get-PfxCertificate -FilePath input.pfx | 
Export-Certificate -FilePath output.der -Type CERT