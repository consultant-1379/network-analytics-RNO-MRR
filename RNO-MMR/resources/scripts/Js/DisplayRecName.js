groupCount = $("#groupCount").text().trim()
fileCount = $("#fileCount").text().trim()
loadedTable = $("#loadedTable").text().trim()

if(loadedTable=='single'){
if(fileCount == 1){
	$("#fileSelected").css("display","")
	$("#groupSelected").css("display","none")
}else if(fileCount == 2){
	$("#fileSelectedComparison").css("display","")
	$("#groupSelectedComparison").css("display","none")
}}
else if(loadedTable=='group'){
if(groupCount == 1){
	$("#fileSelected").css("display","none")
	$("#groupSelected").css("display","")
}else if(groupCount == 2){
	$("#fileSelectedComparison").css("display","none")
	$("#groupSelectedComparison").css("display","")
}}