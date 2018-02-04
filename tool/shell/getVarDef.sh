cat aa|awk ' {
		if ($0 ~ /typedef/) {
			gsub(/\*/,"");
			split($0, fld, "/");
			gsub(/^[ \t]+/, "", fld[2]);
			gsub(/[ \t]+$/, "", fld[2]);
			print "struct",fld[2];
		}
		else if ($0 !~ /}/) {
			split($0, fld, ";");
			split(fld[1], var);
			gsub(/\*/, "", fld[2]);
			gsub(/^[ \t]+/, "", fld[2]);
			gsub(/[ \t]+$/, "", fld[2]);
			split(fld[2], comment, "/");
			gsub(/^[ \t]+/, "", comment[2]);
			gsub(/[ \t]+$/, "", comment[2]);
			
			arrC = 0;
			varname=var[2];
			if ($0 ~ /\[/) {
				str_p=index($0, "[")+1;
				str_len=index($0, "]") - str_p;
				arrC = substr($0, str_p, str_len);
				split(var[2],fld2, "[");
				varname=fld2[1];
			}
			var[2]
			printf("%s,%s,%d,%s\n", var[1],
				varname,arrC,comment[2]);
		}
	}'
