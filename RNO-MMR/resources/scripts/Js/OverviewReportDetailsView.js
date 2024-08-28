fileCount = $("#fileCount").text().trim()
groupCount = $("#groupCount").text().trim()
fileCount = parseInt(fileCount)
groupCount = parseInt(groupCount)

if(fileCount == 1 && groupCount == 0){
$("#comparison-button").css("display","none")
//$("#load-data").css("display","block")
}