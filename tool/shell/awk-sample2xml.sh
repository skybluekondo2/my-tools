cat cc|awk '
	BEGIN {
		f_flg=0;
		print "<?xml version=\"1.0\" encoding=\"shift_jis\"?>";
		print "<awk-samples>"
	}
	{
		#//gsub(/&/,"&amp;");
		#//gsub(/</,"\\&lt;");
		#//gsub(/>/,"\\&gt;");
		if ($0 ~ /#--/) {
			if (f_flg == 1) {
				endTag_code();
				print "</sample>";
			}
			gsub(/#--/,"");
			gsub(/^ +/,"");
			gsub(/ +$/,"");
			print "<sample>";
			printf("<title>%s</title>\n",$0);
			print "<code>";
			print "<![CDATA[";
			f_flg=1;
		} 
		else {
			print $0;
		}
	}
	END {
		endTag_code();
		print "</sample>";
		print "</awk-samples>";
	}
	function endTag_code()
	{
		print "]]>";
		print "</code>";
	}
	'
