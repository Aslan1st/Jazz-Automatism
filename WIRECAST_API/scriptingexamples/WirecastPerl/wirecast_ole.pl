#!/usr/bin/perl

# use ppm to install Win32::OLE as follows:
# 	ppm install win32-ole
use Win32::OLE;

#example of getting the wirecast object (either get current active, or create a new):
my $wirecast = Win32::OLE->GetActiveObject("Wirecast.Application");
if ($wirecast == 0) {
  $wirecast = Win32::OLE->new("Wirecast.Application");
}

if ($wirecast) {
  
  # an example of calling a method:
  my $doc = $wirecast->DocumentByIndex(1);
  
  if ($doc) {
    
    #an example of getting a property:
    printf "TransitionSpeed is %s\n", $doc->{TransitionSpeed};
    
    #an example of setting a property:
    $doc->{TransitionSpeed} = "fast";	
    
    #verify property was set:
    printf "TransitionSpeed is %s\n", $doc->{TransitionSpeed};
  }
}

print "done.\n";
