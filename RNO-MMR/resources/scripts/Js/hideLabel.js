MutationObserver = window.MutationObserver || window.WebKitMutationObserver;
var changeUI = function(){
var box = $("#cell").text()
console.log(box)
var bx_len = box.length
if(bx_len != 0){

	$("#cell").hide();
}
    else{
		$("#cell").hide();
		

	}
       
}

var target = document.getElementById("cell")

//callback is the function to trigger when target changes
var callback = function(mutations) {
    changeUI()
}

var observer = new MutationObserver(callback);
var opts = {
    childList: true, 
    attributes: true, 
    characterData: true, 
    subtree: true
}
observer.observe(target,opts);
changeUI()