#!/bin/bash
cat $1|gawk '/#define/ {
	gsub(/:/,"");
	gsub(/\*/,"");
	split($0, fld, "/");
	gsub(/^[ \t]+/, "", fld[2]);
	gsub(/[ \t]+$/, "", fld[2]);
	printf("%s,%s,%s,%s\n",$2,$3,fld[2],category);
	}
	/^\// {
		gsub(/\*/,"");
		split($0, fld, "/");
		gsub(/^[ \t]+/, "", fld[2]);
		gsub(/[ \t]+$/, "", fld[2]);
		category=fld[2];
	
	}
'
