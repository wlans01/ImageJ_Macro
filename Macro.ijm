// data 파일이 있는 절대경로 기입
// data 파일에는 처리를 하고싶은 이미지만 존재해야함
//ex) "C:\\Users\\User\\Desktop\\macro\\data\\"

images_path ="C:\\Users\\User\\Desktop\\macro\\data\\"

// temp 파일이 있는 절대 경로 기입
// ex) 'C:\\Users\\User\\Desktop\\macro\\temp\\'

result_path ='C:\\Users\\User\\Desktop\\macro\\temp\\'

//data에서 파일 찾기
folderList = getFileList(images_path);


// 이미지를 처리하고 xls로 저장
for (i=0; i<folderList.length; i++){
    fileList=getFileList(images_path+folderList[i]);
    for (j=0; j<fileList.length; j++){
        open(images_path+folderList[i]+fileList[j]);
        run("Select All");
        run("Measure");
        //setTool("line");
        makeLine(0, 1458, 1936, 0);
        run("Measure");
        saveAs( "Results",result_path+folderList[i]+fileList[j]+".xls");
        String.copyResults();
        IJ.deleteRows(0, 1);
        run("Close All");
}}





