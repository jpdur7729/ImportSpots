rem It works without adding parameters which is probably good enough in most circumstances
wget "http://data.fixer.io/api/2013-12-24?access_key=eca17521f4e211d09ab357c6cd9585dc" -o LogFile.txt
                                                                                          
rem ------------------------------------------------------------------------
rem In .bat file you have to add a && if not & is not interpreted correctly 
rem http://www.robvanderwoude.com/escapechars.php
rem ------------------------------------------------------------------------
rem Worked in the Browser http://data.fixer.io/api/2013-12-24?access_key=eca17521f4e211d09ab357c6cd9585dc&base=EUR&symbols=USD,CAD,EUR
wget "http://data.fixer.io/api/2013-12-24?access_key=eca17521f4e211d09ab357c6cd9585dc&&base=EUR&&symbols=USD,CAD,EUR" -o LogFile.txt
