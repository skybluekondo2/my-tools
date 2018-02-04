function createBook(sFile) {
	// Excelオブジェクトを作成する。
	var excel = new ActiveXObject("Excel.Application");

	// 可視プロパティを可視(=true)に設定する。
	excel.Visible = true;

	// 警告表示プロパティを非表示(=false)に設定する。
	excel.DisplayAlerts = false;
	try {
	    // ワークブックを作成する。
			var workbook = excel.Workbooks.Add();
			workbook.SaveAs(sFile);
	} finally {
	    excel.Quit();
	}
}

// WshShell オブジェクトを作成する。
var WshShell = WScript.CreateObject("WScript.Shell");

// デスクトップフォルダを取得する。
strDesktop = WshShell.SpecialFolders("Desktop");

// デスクトップへ新規ブックを作成する。
createBook(strDesktop + "/newBook.xls");
