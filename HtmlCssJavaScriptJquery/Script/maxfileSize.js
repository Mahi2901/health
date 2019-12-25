var attachementSizeInKB=0;
function checkFileSize(attachement) {
	try {
		var attachementSize=0;
		//var attachement = document.getElementById('attachementLotFileID');
		var file = attachement.files[0];
		var bytes = file.size;
		var sizes = [ 'Bytes', 'KB', 'MB', 'GB', 'TB' ];
		
		if (bytes == 0)
			return 'n/a';
		var i = parseInt(Math.floor(Math.log(bytes) / Math.log(1024)));
		if (sizes[i] == "MB") {
			attachementSize = Math.round(bytes / Math.pow(1024, i), 2);
		} 
		else if(sizes[i] == "KB") {
			attachementSizeInKB = Math.round(bytes / Math.pow(1024, i), 2);
		}
		else if(sizes[i] == "Bytes") {
			attachementSizeInKB = bytes;
		}
		else {
			attachementSize = 0;
		}
	} catch (e) {
		attachementSize = 0;
	}
	return attachementSize;
}
