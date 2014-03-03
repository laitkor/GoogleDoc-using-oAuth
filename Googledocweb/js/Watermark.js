 
        function WatermarkFocus(txtElem, strWatermark) {
            if (txtElem.value == strWatermark) txtElem.value = '';
        }

        function WatermarkBlur(txtElem, strWatermark) {
            if (txtElem.value == '') txtElem.value = strWatermark;
        }

        function isNumberKey(evt) {
            var charCode = (evt.which) ? evt.which : event.keyCode
            if (charCode > 31 && (charCode < 48 || charCode > 57))
                return false;
               return true;
           }

           function NotAllow() {
               var iChars = "!@#$%^&*()+=-[]\\\';,./{}|\":<>? ";
               for (var i = 0; i < document.form1.ContentPlaceHolder1_txtUser.value.length; i++) {
                   if (iChars.indexOf(document.form1.ContentPlaceHolder1_txtUser.value.charAt(i)) != -1) {
                       alert("Your username has special characters. \nThese are not allowed.\n Please remove them and try again.");
                       return false;
                   }
               }
           }
    