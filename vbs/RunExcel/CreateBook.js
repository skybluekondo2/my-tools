function createBook(sFile) {
	// Excel�I�u�W�F�N�g���쐬����B
	var excel = new ActiveXObject("Excel.Application");

	// ���v���p�e�B����(=true)�ɐݒ肷��B
	excel.Visible = true;

	// �x���\���v���p�e�B���\��(=false)�ɐݒ肷��B
	excel.DisplayAlerts = false;
	try {
	    // ���[�N�u�b�N���쐬����B
			var workbook = excel.Workbooks.Add();
			workbook.SaveAs(sFile);
	} finally {
	    excel.Quit();
	}
}

// WshShell �I�u�W�F�N�g���쐬����B
var WshShell = WScript.CreateObject("WScript.Shell");

// �f�X�N�g�b�v�t�H���_���擾����B
strDesktop = WshShell.SpecialFolders("Desktop");

// �f�X�N�g�b�v�֐V�K�u�b�N���쐬����B
createBook(strDesktop + "/newBook.xls");
