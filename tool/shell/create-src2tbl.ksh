cat *.*|awk '
	/^static.+\[\].+\{/ {
		print;
		k_flg=1;
		while (getline) {
			tmp=$0;
			gsub(/\t/, "", tmp);
			if (tmp !~ / +\/\*/) {
				print;
			}
			if ($0 ~ /\{/) k_flg++;
			if ($0 ~ /\}/) k_flg--;
			if ($0 ~ /\}/ && k_flg == 0) break;
		}
	}
	'
