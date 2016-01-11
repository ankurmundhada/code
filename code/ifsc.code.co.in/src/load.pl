#!/usr/bin/perl

# PERL MODULES WE WILL BE USING
use DBI;
use DBD::mysql;
use WWW::Mechanize;
use File::Fetch;
use Spreadsheet::XLSX;
use Spreadsheet::ParseExcel;

# CONFIG VARIABLES
$platform = "mysql";
$database = "code";
$host = "localhost";
$port = "3306";
$user = "root";
$pw = "";

# DATA SOURCE NAME
$dsn = "dbi:$platform:$database:$host:$port";

# PERL DBI CONNECT
$connect = DBI->connect($dsn, $user, $pw);

my $i=1;
my $flag;
my $filename;
my $err='';
my $sum=0;
my $pflag = 0;

my @url;
my $mech = WWW::Mechanize->new();
$mech->get( 'http://www.rbi.org.in/Scripts/bs_viewcontent.aspx?Id=2009' );
my @links = $mech->links();
for my $link ( @links ) {
	my $links = $link->url;
	if ($links =~ m/xls/) {
		if ( $links =~ m/_122/ || $links =~ m/_48/ )
		{	#printf "%s, %s\n", $link->text, $link->url;
			push (@url,$links);			
		}
	}
}
#print join("\n",@url);
#exit;
foreach $link (@url)
{
	@ar = split (/\//, $link);
	$filename = @ar[$#ar];
	
	# to download file
	#if ( $filename =~ m/_122/ || $filename =~ m/_48/ )
	#{ } else { 
		my $url = $link;
		my $ff = File::Fetch->new(uri => $url);
		my $file = $ff->fetch() ;#or die $ff->error;
		$err = $ff->error;
		
		if ($err eq "")
		{	
			print " \n Downloaded  $url \n";
			$flag = 1;
		}else
		{
			print "\n NOT Downloaded $url \n";
			$flag = 0;
		}
	#}	
	sleep(2);	
	
	if ( $filename eq 'IFCB2009_101.xls') {next;}
	if ( $filename eq 'IFCB2009_98.xls') {next;}
	
	#$filename = 'IFCB2009_'.$i.'.xls';
	print 'For File:-'.$filename."\n";
	eval 
	{
			my $excel = Spreadsheet::XLSX -> new ($filename) or die 'some error';
			foreach my $sheet (@{$excel -> {Worksheet}}) {
				#printf("Sheet: %s\n", $sheet->{Name});
				$sheet -> {MaxRow} ||= $sheet -> {MinRow};
				foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}) {
					if ($row == 0 ) {next;}	
					$sheet -> {MaxCol} ||= $sheet -> {MinCol};
					foreach my $col ($sheet -> {MinCol} ..  $sheet -> {MaxCol}) {
					my $cell = $sheet -> {Cells} [$row] [$col];
					$coldata = $cell -> {Val} ;
						if ($cell) {
							 $coldata =~ s/"//g;
							 $coldata =~ s/'//g;
							 $coldata =~ s/[^[:ascii:]]//g;
							 $coldata =~ s/\f//g;
							 $coldata =~ s/&#225;//g;
							 $coldata =~ s/\r//g;
							 $coldata =~ s/\n//g;
							#printf("( %s , %s ) => %s\n", $row, $col, $cell -> {Val});
							if ($col == 0 ) { $BANK = $coldata };
							if ($col == 1 ) { $IFSC_CODE = $coldata };
							if ($col == 2 ) { $MICR_CODE = $coldata };
							if ($col == 3 ) { $BRANCH_NAME = $coldata };
							if ($col == 4 ) { $ADDRESS = $coldata };
							if ($col == 5 ) { $CONTACT = $coldata };
							if ($col == 6 ) { $CITY = $coldata };
							if ($col == 7 ) { $DISTRICT = $coldata };
							if ($col == 8 ) { $STATE = $coldata };					
						}
					}
					if ($IFSC_CODE ne ""){
						$query = "INSERT INTO IFSC (IFSC_CODE, MIRC_CODE, BANK, BRANCH, ADDRESS, CONTACT, CITY, DISTRICT, STATE) 
						VALUES ('$IFSC_CODE','$MICR_CODE','$BANK','$BRANCH_NAME','$ADDRESS','$CONTACT','$CITY','$DISTRICT','$STATE')";
						$query_handle = $connect->prepare($query);
						$query_handle->execute();	
						$sum++;
					}
					
					$IFSC_CODE = "";
				}				
			}  # -- end Spreadsheet::XLSX			
			print "\n DONE WITH MODULE 1";
			print " \n Compleated Loading $sum";
	};	
	if ($@){
		print "\n Using Differant module to parse file \n";
		$usemod2 = 1;	
    }
	else{
		#print " \n Compleated Loading $sum";
	}
	
	if ($usemod2 == 1)
	{
		my $source_excel = new Spreadsheet::ParseExcel;
		my $source_book = $source_excel->Parse($filename) ;		
		foreach my $source_sheet_number (0 .. $source_book->{SheetCount}-1)
		{
		 my $source_sheet = $source_book->{Worksheet}[$source_sheet_number];
		 #print $source_sheet->{Name};
		 next unless defined $source_sheet->{MaxRow};
		 next unless $source_sheet->{MinRow} <= $source_sheet->{MaxRow};
		 next unless defined $source_sheet->{MaxCol};
		 next unless $source_sheet->{MinCol} <= $source_sheet->{MaxCol};

		 foreach my $row_index ($source_sheet->{MinRow} .. $source_sheet->{MaxRow})
		 {
	
					  if ($row_index == 0 ) {next;}	
					  my $line = '';	
					  foreach my $col_index ($source_sheet->{MinCol} .. $source_sheet->{MaxCol})
					  {		  
	
					   my $source_cell = $source_sheet->{Cells}[$row_index][$col_index];
					   if ($source_cell)
					   {
							$coldata = $source_cell->Value;
							if ($coldata ne "")
							{
								 $coldata =~ s/"//g;
								 $coldata =~ s/'//g;
								 $coldata =~ s/[^[:ascii:]]//g;
								 $coldata =~ s/\f//g;
								 $coldata =~ s/&#225;//g;
								 $coldata =~ s/\r//g;
								 $coldata =~ s/\n//g;
								#print "( $row_index , $col_index ) =>", $source_cell->Value, "\t";
								if ($col_index == 0 ) { $BANK1 = $coldata };
								if ($col_index == 1 ) { $IFSC_CODE1 = $coldata };
								if ($col_index == 2 ) { $MICR_CODE1 = $coldata };
								if ($col_index == 3 ) { $BRANCH_NAME1 = $coldata };
								if ($col_index == 4 ) { $ADDRESS1 = $coldata };
								if ($col_index == 5 ) { $CONTACT1 = $coldata };
								if ($col_index == 6 ) { $CITY1 = $coldata };
								if ($col_index == 7 ) { $DISTRICT1 = $coldata };
								if ($col_index == 8 ) { $STATE1 = $coldata };
							}
						}
					  } 
					if ($IFSC_CODE1 ne "") {						
						$query = "INSERT INTO IFSC (IFSC_CODE, MIRC_CODE, BANK, BRANCH, ADDRESS, CONTACT, CITY, DISTRICT, STATE) 
						VALUES ('$IFSC_CODE1','$MICR_CODE1','$BANK1','$BRANCH_NAME1','$ADDRESS1','$CONTACT1','$CITY1','$DISTRICT1','$STATE1')";
						$query_handle = $connect->prepare($query);
						$query_handle->execute();	
						$sum++;
					}
					$IFSC_CODE1 = "";
			}		 
		}
		print "\n DONE WITH MODULE 2";	
		print " \n Compleated Loading $sum";
	}
	$usemod2 = 0;
}
