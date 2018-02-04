if [ $# -ne 3 ];then
	echo "Usage:`basename $0` [file] [str1] [str2]";
	exit 1;
fi

L_FILE=$1;
L_STR_FROM=$2;
L_STR_TO=$3;

cat ${L_FILE}|awk -F, '!/struct/ {
	printf("\tmemcpy(\t\t\t/* %-16s\t*/\n",$4);
	printf("\t\t%s->%s,\n", "'${L_STR_TO}'",$2);
	printf("\t\t%s_p->%s,\n", "'${L_STR_FROM}'",$2);
	printf("\t\tsizeof(%s->%s));\n", "'${L_STR_TO}'",$2);
	}'
