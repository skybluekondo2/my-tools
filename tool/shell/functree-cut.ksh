#!/usr/bin/bash
#-- function tree Œ‹‰ÊØ‚èo‚µ
for x in `ls *.txt`
do
	echo "--------- `basename $x ".txt"|sed 's/ftree-//g'`() ----------"
	cat $x|awk '
		/()/ {
			for(i=1;i<=NF;i++) {
				if ($i ~ /</) {
					lv=$i;
					gsub(/</,",",lv);
					gsub(/>/,",",lv);
					split(lv,fld,",");
					for(j=1;j<=fld[2];j++) {
						printf("    ");
					}
					print $i;
				}
			}
		}'
done
