this.LargeImagePreview = function () {

	// these 2 variable determine popup's distance from the cursor
	// you might want to adjust to get the right result
	xOffset = 10;
	yOffset = 10;

	// View Large Image on Mouse Hover Event
	$("a.ImagePreview").hover(function (e) {
		//Get rel Data from hidden Image control.
		var ImgHidden = $(this).attr('rel');

		// Change String 
		// For Example - ~/Images/Kakashi.jpg to Images/Kakashi.jpg
		var ImgSrc = ImgHidden.replace("~/", "");

		// Bind Images in Paragraph tag
		$("body").append("<p id='ImagePreview'><img src='" + ImgSrc + "' alt='loading...' /></p>");

		$("#ImagePreview")
			.css("top", (e.pageY - xOffset) + "px")
			.css("left", (e.pageX + yOffset) + "px")
			.fadeIn("fast");
	},
	function () {
		$("#ImagePreview").remove();
	});
$("a.ImagePreview").mousemove(function (e) {
		$("#ImagePreview")
			.css("top", (e.pageY - xOffset) + "px")
			.css("left", (e.pageX + yOffset) + "px");
	});

};