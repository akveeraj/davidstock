#!/usr/bin/perl
##
## $Version$

$log_file="/sites/davidostick.co.uk/hitslogger/hits.html";

print "Content-type: text/html\n\n";

$date=gmtime(time);

# Suck up the old file
open (FILE,"< $log_file") or die "Can't read $log_file!";
@lines=<FILE>;
close FILE;

open (LOG,">$log_file") or die "Can't write to $log_file!";
foreach (@lines) {   
	s/<!--edit-->/<<PAGE/e;
<!--edit-->
<tr><td>
<table style='width:100%'>
	<tbody style='background-color: white'>
		<tr>
			<td style='width:10%'>
				<b>Date:</b>
			</td><td style='width:90%'>
				$date
			</td>
		</tr><tr>
			<td>
				<b>IP Address:</b>
			</td><td>
				$ENV{REMOTE_ADDR}
			</td>
		</tr><tr>
			<td>
				<b>HostName:</b>
			</td><td>
				$ENV{REMOTE_HOST}
			</td>
		</tr><tr>
			<td>
				<b>Browser/OS:</b>
			</td><td>
				$ENV{HTTP_USER_AGENT}
			</td>
		</tr><tr>
			<td>
				<b>Came from:</b>
			</td><td>
				<a href="$ENV{HTTP_REFERER}">$ENV{HTTP_REFERER}</a>
			</td>
		</tr><tr>
			<td>
				<b>Query:</b>
			</td><td>
				$ENV{QUERY_STRING}
			</td>
		</tr>
	</tbody>
</table>
</td></tr>
PAGE
	print LOG $_; 
}   
close LOG;

exit 0;