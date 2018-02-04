#!/usr/bin/bash
cat $1|awk '
	/#def/ {
		gsub(/\*/,"");
		split($0, fld, "/");
		gsub(/^[ \t]+/,"",fld[2]);
		gsub(/[ \t]+$/,"",fld[2]);
		printf("%s,%s,\"%s\"\n",$3,$2,fld[2]);
	}
	'
