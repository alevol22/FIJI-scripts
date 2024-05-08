#@ File (label = "Input directory", style = "directory") input
#@ File (label = "Output directory", style = "directory") output

processFolder(input, output);

// function to scan folders/subfolders/files to find files with correct suffix
function processFolder(input, output) {
	list = getFileList(input);
	LD = newArray(list.length);
	list = Array.sort(list);
	files = newArray(list.length);
	j = 0;
	
	for (i = 0; i < list.length; i++) {
		if(endsWith(list[i], ".tif")){
			file = list[i];
			open(input + File.separator + file);
			run("Subtract Background...", "rolling=2.5");
			fileName = substring(file,0,lengthOf(file)-8);
			print(fileName);
			saveAs("Tiff" , output + File.separator + fileName + "_RB2p5.tif");
  			close();
		}
		
	}

}