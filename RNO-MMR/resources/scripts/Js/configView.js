fileCount = $("#fileCount").text().trim()
groupCount = $("#groupCount").text().trim()

fileCount = parseInt(fileCount)
groupCount = parseInt(groupCount)

fileCount2 = $("#fileCount2").text().trim()
groupCount2 = $("#groupCount2").text().trim()

if(fileCount == 1){
	$("#comparison-button").css("display","none")
	$("#single-recording").css("display","")
	$("#group-recording").css("display","none")
	$("#export-button").css("display","")
	$("#details-buttons").css("display","")
	$("#trends-buttons").css("display","none")
}else if(fileCount == 2){
	$("#comparison-button").css("display","")
	$("#single-recording").css("display","none")
	$("#group-recording").css("display","none")
	$("#export-button").css("display","none")
	$("#details-buttons").css("display","none")
	$("#trends-buttons").css("display","")
}else if(fileCount >= 3){
	$("#comparison-button").css("display","none")
	$("#details-buttons").css("display","none")
	$("#trends-buttons").css("display","")
	$("#export-button").css("display","none")
}else{
	$("#comparison-button").css("display","none")
	$("#removeGrouping").css("display","")
	$("#single-recording").css("display","none")
	$("#group-recording").css("display","none")
	$("#details-buttons").css("display","none")
	$("#export-button").css("display","none")
	$("#fetch-button").css("display","none")
	//$("#trends-buttons").css("display","block")
}
if(groupCount == 1){
	$("#comparison-button-group").css("display","none")
	$("#single-recording-group").css("display","")
	$("#group-recording-group").css("display","none")
	$("#export-button-group").css("display","")
	$("#details-buttons-group").css("display","")
	$("#trends-buttons-group").css("display","none")
}else if(groupCount == 2){
	$("#comparison-button-group").css("display","")
	$("#single-recording-group").css("display","none")
	$("#group-recording-group").css("display","none")
	$("#export-button-group").css("display","none")
	$("#details-buttons-group").css("display","none")
	$("#trends-buttons-group").css("display","")
}else if(groupCount >= 3){
	$("#comparison-button-group").css("display","none")
	$("#details-buttons-group").css("display","none")
	$("#trends-buttons-group").css("display","")
	$("#export-button-group").css("display","none")
}else{
	$("#comparison-button-group").css("display","none")
	$("#removeGrouping-group").css("display","none")
	$("#single-recording-group").css("display","none")
	$("#group-recording-group").css("display","none")
	$("#details-buttons-group").css("display","none")
	$("#export-button-group").css("display","none")
	$("#fetch-button-group").css("display","none")
	//$("#trends-buttons").css("display","block")
}

/*
fileCount = parseInt(fileCount)
groupCount = parseInt(groupCount)
if(fileCount2 == 0){
$("#load-data").css("display","none")
}
else{
	$("#load-data").css("display","")
}

if(groupCount2 == 0){
	$("#load-data-group").css("display","none")
	$("#removeGrouping").css("display","")
}
else{
	$("#load-data-group").css("display","")
}
console.log("fileCount2 == "+fileCount2)
console.log("groupCount2 == "+groupCount2)

*/