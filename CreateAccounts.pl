# Script takes in an enrollment list
# scans AD and compares to list passed in
# if they're not in AD, it creates them, assigns appropriate group membership, and puts in correct OU
# Author: Barret J. Nobel
# Contact: bear.nobel@gmail.com
# UPDATED: 12/16/15

use Win32::OLE;
use strict;
use warnings;
use 5.10.0;
use Text::CSV;
use Win32::OLE 'in';
use Net::LDAP;

my $enrollmentList; #................................................... variable to hold the enrollment list file passed in
my $csv = Text::CSV->new({sep_char => ','}); #.......................... parser used is a comma ( , )
my $outputNamesFile = "[REDACTED]/namesNotInAD.txt"; #....... variable to access the output file for the names not in AD
my $successAddedFile = "[REDACTED]/successfullyAdded.txt"; #. file to print out successfully created names
my $ouErrorFile = "[REDACTED]/OU_ERROR.txt"; #............... file that is created when the OU is not found
my @peopleNotInAD; #.................................................... array to hold the people objects who're not in AD

###############################################
#               DONT EDIT [REDACTED]                #
###############################################
####### This section connects up to AD ########
my $dse=Win32::OLE->GetObject("LDAP://RootDSE");
my $root=$dse->Get("RootDomainNamingContext");
my $adpath="GC://$root";
my $base="<".$adpath.">";
my $connection = Win32::OLE->new("ADODB.Connection");
   $connection->{Provider} = "ADsDSOObject";
   $connection->Open("ADSI Provider");
my $command=Win32::OLE->new("ADODB.Command");
   $command->{ActiveConnection}=$connection;
   $command->{Properties}->{'Page Size'}=1000;
my $rs = Win32::OLE->new("ADODB.RecordSet");
###############################################
#               DONT EDIT [REDACTED]                #
###############################################

&evaluateArgs(); #............................................................................................. checks to see if the argument passed in is correct
open(CSV, '<', $enrollmentList) or die "Error opening file: $enrollmentList.\n$!\n"; #......................... opens up the enrollment list passed in 
open(my $nameOutputStream, '>', $outputNamesFile) or die "Error opening file: $outputNamesFile.\n$!\n"; #...... opens up the file to write names not in AD
open(my $successOutputStream, '>', $successAddedFile) or die "Error opening file: $successAddedFile.\n$!\n"; #. opens up the file to write successfully added accounts
&loopThroughEnrollmentList(); #................................................................................ loops through enrollment list ( incase you couldnt tell by the name :) )
&createUserInAD(); #........................................................................................... creates users in AD that were found in the loopThroughEnrollmentList method above
   
##### CLOSES ALL FILES #####
close CSV;
close $nameOutputStream;
close $successOutputStream;

# Method to create a user in AD
# uses array @peopleNotInAD
# parses and creates the user in AD
# Info in Array:
#  0: Student ID number
#  1: User Name
#  2: ?
#  3: First and last name
#  4: Last name
#  5: First name
#  6: Middle initial
#  7: User Status
#  8: Address
#  9: Email
# 10: Phone
# 11: Password
# 12: Location
# 13: AD Location
sub createUserInAD
{	
	print $successOutputStream "@{[get_timestamp(1)]}\n\n"; #............................................................. print the headers to the file successAddedFile.txt
	printf $successOutputStream "%-20s%-20s%-10s\n\n", "First", "Last", "Username";
	my $numberOfSuccessfullyCreatedUsers = 0; #........................................................................... variable to count the number of users successfully added
	
	foreach( @peopleNotInAD ) #........................................................................................... loop through each person not in AD
	{	
		print "\nCreating users in AD now ... \n\n";
		
		my @arrayTest = $_; #............................................................................................. get the first person ( they're arrays )
		
		my $sevenHundredNum = $arrayTest[0][0]; #......................................................................... parse the array and assign it to variables ( Might be better as a hash? )
		my $samAccountName  = $arrayTest[0][1]; 
		my $displayName     = $arrayTest[0][1];
		my $principalName   = $arrayTest[0][1];
		my $lastName        = $arrayTest[0][4];
		my $firstName       = $arrayTest[0][5];
		my $initial         = $arrayTest[0][6];
		my $facOrEmp        = $arrayTest[0][7];
		my $emailAddress    = $arrayTest[0][9];
		my $password        = $arrayTest[0][11];
		my $ouFromArray    	= $arrayTest[0][13];
		my $cnName         	= "$lastName " . $initial . " $firstName";
		
		my $ouString = &createOUString( $ouFromArray, $facOrEmp, $firstName, $lastName, $sevenHundredNum ); #............. create the OU string
		
		print "Creating $firstName: ";
		my $ou=Win32::OLE->GetObject("LDAP://ou=$ouString,dc=[domain path]") or die Win32::OLE->LastError();; #...... connects to the correct OU based on the string created above
		my $user=$ou->Create("User","cn=$cnName") or die "Unable to create user $firstName\n"; #.......................... creates the user in the OU
		
		sleep(3);
		
		   $user->{sAMAccountName}    = $samAccountName; #................................................................ assign object properties
		   $user->{displayName} 	  = $displayName;
		   $user->{name}         	  = $firstName; 
		   $user->{givenName}         = $firstName; 
		   $user->{sn}                = $lastName;
		   $user->{mail}              = $emailAddress;
		   $user->{initials}          = $initial if( $initial );
		   $user->{userPrincipalName} = $principalName . '@[REDACTED].edu';
		   $user->SetInfo( ); #........................................................................................... method needed to 'save' commands assigned to the object
	   
		   $user->{userAccountControl} = '512'; #......................................................................... unlocks the account
		   $user->SetInfo( );
		   
		   $user->ChangePassword("",$password); #......................................................................... sets password to uncheck the box User must Change on next logon
		   $user->SetInfo( );
		 
		print "Success!\n";
		   
		print "\nAssigning group membership for $firstName\n";
		&assignGroupMembership( $ouString, $cnName ); #................................................................... assign membership
		print "Successfully assigned\n\n";
		
		printf $successOutputStream "%-20s%-20s%-10s\n", $firstName, $lastName, $samAccountName; #........................... prints to a file the names of people successfully added
		$numberOfSuccessfullyCreatedUsers ++;
	}
	
	printf $successOutputStream "\n\nThere were %d successfully created accounts", $numberOfSuccessfullyCreatedUsers; #...... outputs the number of successfully created users
}

# Method to assign the user the appropriate group membership
# you've to add the user to the group, NOT add the group to the user object
# takes two arguments
# 1: full OU Path ( String )
# 2: cn name
sub assignGroupMembership()
{	
	say "Sleeping for 5 seconds before assigning membership\n"; #.......................................................................... sleeps to give AD time to create account
	sleep (5);
	my $fullOuString  = $_[0]; #........................................................................................................... capture data passed in
	my $cnName        = $_[1];
	my $cnPath; #.......................................................................................................................... variable to hold cnPath
	
	my $domain = ",dc=[domain path]"; #............................................................................................... variable to hold domain name
	my @ouPathValues = split( ',', $fullOuString ); #...................................................................................... parses the fullOuString passed in
	
	my @cnNames = ( '[REDACTED]' ); #........................................................................................ array full of exceptions for the cnPath
	
	if( $ouPathValues[0] eq "[REDACTED]\\" ) { $cnPath = "[REDACTED]"; } #.......................................................................... special case when the OU has commas in the name
	else
	{
		$cnPath = ($ouPathValues[0] ~~ @cnNames) ? &createCNPath( $ouPathValues[0] ) : $ouPathValues[0]; #................................. get cnPath from the parsed full path ( checks to see if its in the array above first )
	}
	
	my $groupObject = Win32::OLE->GetObject('LDAP://cn=' . $cnPath . ',ou=' . $fullOuString . $domain) or die Win32::OLE->LastError(); #... gets the group from AD based on path
       $groupObject->Add('LDAP://cn=' . $cnName . ',ou=' . $fullOuString . $domain); #..................................................... add the user to the group
	Win32::OLE->LastError(); #............................................................................................................. if there's an error
}

# Method to create the CN path if it's different than the OU name
# takes 1 argument
# the path value parsed from the OUString in the assignGroupMembership() method
sub createCNPath
{
	my $pathValue = $_[0]; #................................................. argument passed in
	my $returnString; #...................................................... variable to hold the string to be returned
	
	given( $pathValue ) #.................................................... based on value passed in, assign the return string
	{
		when( "[REDACTED]" ) 	{ $returnString = "[REDACTED]"; }
		when( "[REDACTED]" ) 	{ $returnString = "[REDACTED]"; }
		when( "[REDACTED]" ) 	{ $returnString = "[REDACTED]"; }
		when( "[REDACTED]" ) 	{ $returnString = "[REDACTED]"; }
		when( "[REDACTED]" ) 	{ $returnString = "[REDACTED]"; }
	}
		
	return $pathValue;
}

# Method to create the string for the OU
# takes in five arguments
# 1: OU Name ( String )
# 2: faculty or employee ( String )
# 3: First name
# 4: Last name
# 5: 700 number
# returns a string or throws an error 
sub createOUString
{
	my $ou     			 = $_[0]; #................. the OU passed in
	my $fOrE   			 = $_[1]; #................. facutly or employee variable passed in
	my $fName  			 = $_[2]; #................. First name
	my $lName  			 = $_[3]; #................. Last name
	my $sevenHundredNum  = $_[4]; #................. Student ID Number #
	my $returnString = ""; #........................ variable to hold the returnString with the OU name
	
	open(my $errorOutputStream, '>', $ouErrorFile) or die "Error opening file: $ouErrorFile.\n$!\n"; #..... opens and sets up the file to print errors to
	
	if( $fOrE eq "employee" ) #..................... if its an employee
	{
		given( $ou )
		{
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
		
			default { 
						print "************Error, check log***************\n";
						printf $errorOutputStream "OU %s not found for user\n%-7s %s\n%-7s %s\n%-7s %s\n*********", $ou, "First:", $fName, "Last:", $lName, "700:", $sevenHundredNum; 
					}
		}
	}
	else #.......................... if they're faculty
	{
		given( $ou )
		{
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }
			when ( "[REDACTED]" )	{ $returnString = "[REDACTED], ou=[REDACTED]"; }

		
			default { 
						print "************Error, check log***************\n";
						printf $errorOutputStream "OU %s not found for user\n%-7s %s\n%-7s %s\n%-7s %s\n*********", $ou, "First:", $fName, "Last:", $lName, "700:", $sevenHundredNum; 
					}
		}
	}
	
	close $errorOutputStream;
	return $returnString;
}

# Method that searches AD for the username passed in
# Currently returns the name if its found
# Need to update to throw an error
sub searchADForUser
{
	my $accountNameToSearch = $_[0]; #........................................................................................ argument passed in ( the username )
	my $returnName; #......................................................................................................... variable to store the name to return

	$command->{CommandText}="$base;(&(objectCategory=Person)(sAMAccountName=$accountNameToSearch));displayName;subtree"; #.... command to get the AD object 
	$rs=$command->Execute;
	
	until ($rs->EOF) #........................................................................................................ loop through all records
	{
		$returnName=$rs->Fields(0)->{Value}; #................................................................................ assigns returnName the name that doesnt exist
		$rs->MoveNext;
	}
	
	return $returnName;
}

# Method that loops through the enrollment list file passed in
# queries AD to see if they exist
# prints list of users that are not in AD to a file ( namesNotInAD.txt )
# throws all people who don't exist into an array to be added
# format: First Name, Last Name, Department
# ALSO
# [REDACTED]
sub loopThroughEnrollmentList
{
	print "\nChecking enrollment list\n";
	
	my $nameToCheck; #................................................................................................ variables
	my $lastName;
	my $firstName;
	my $department;
	my $userName;
	my $pinNumber;
	
	print $nameOutputStream "@{[get_timestamp(1)]}\n\n"; #............................................................ print the headers to the file namesNotInAD.txt
	printf $nameOutputStream "%-20s%-20s%-10s\n\n", "First", "Last", "Department";
	
	while(<CSV>) #.................................................................................................... while there is still stuff to loop through
	{
		if($csv->parse($_))
		{
			my @input_array = $csv->fields(); #....................................................................... toss all csv separated fields into an array
			
			$nameToCheck = &searchADForUser($input_array[1]); #....................................................... initialize the name if its not in AD already
			$firstName =  $input_array[5]; #.......................................................................... assign the variables from the array
			$lastName  =  $input_array[4];
			$department = $input_array[13];
			$userName  =  $input_array[1];
			$pinNumber =  $input_array[11];
			my $samName = $input_array[1];
			
			if( not(defined $nameToCheck) && ( &checkExceptionList($samName) == 0) ) #................................ if the name is not defined it doesnt exist! -- MUST ADD TO LIST ( also makes sure its not in the exception list )
			{
				push( @peopleNotInAD, \@input_array ); #.............................................................. passes the array of the person object to the array of people not in AD
				printf $nameOutputStream "%-20s%-20s%-10s\n", $firstName, $lastName, $department; #................... writes names to namesNotInAD.txt file
				[REDACTED]
			}
		}
	}
	
	print [REDACTED]#.............................. prints a temp username for a bug fix if script is empty
	
	print "Done checking enrollment list\n";
}

# Method to check array for exceptions before we create the user
# meaning, they already exist, but they're spelt differently
# NOTE: checks username
# returns 1 if it IS in the list
# returns 0 if its NOT in the list
sub checkExceptionList
{
	my @exceptionList = ( '[REDACTED]' );
	my $nameToCheck = lc $_[0];
	
	return ( $nameToCheck ~~ @exceptionList ) ? 1 : 0;
}

# Method to evaluate the arguments passed in
# throws an error to the console and creates a log
sub evaluateArgs
{
    open(my $log, '>>', "create_users.log"); #.................................................................................. creates / opens log file for errors
    if(scalar(@ARGV) == 1) #.................................................................................................... if script passes in AD list
	{
		$enrollmentList =  $ARGV[0]; #.......................................................................................... assign it to enrollmentList variable
	}
	else
	{
		say 'usage is: [REDACTED].pl <banner report of accepted students>';
		print $log "@{[get_timestamp()]} The command was incorectly used.  Check the script for any incorrect arguments.\n";
		close $log;
		exit 1;
	 }
}

# Method to give the timestamp
sub get_timestamp 
{
	(my $sec,my $min,my $hour,my $mday,my $mon,my $year,my $wday,my $yday,my $isdst) = localtime(time);
	$mon +=1;
	if ($mon < 10) { $mon = "0$mon"; }
	if ($hour < 10) { $hour = "0$hour"; }
	if ($min < 10) { $min = "0$min"; }
	if ($sec < 10) { $sec = "0$sec"; }
	if ($mday < 10) { $mday = "0$mday"; }
	$year+=1900;
	if($_[0])
	{
		return '[' . $year . '_' . $mon . '_' . $mday . ' ' . $hour . ':' . $min . ':' . $sec . ']';
	}
	return $year . '_' . $mon . '_' . $mday . '_' . $hour . '_' . $min . '_' . $sec;
}