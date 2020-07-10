/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

(function () {
	"use strict";
	Office.initialize = function () {
		var currentDate = getActualDate();
		var request = GetItem();
        var envelope = getSoapEnvelope(request);
		var mBox = Office.context.mailbox;
		var subject = mBox.item.subject;
		mBox.makeEwsRequestAsync(envelope, function(result){
			var parser = new DOMParser();
			var doc = parser.parseFromString(result.value, "text/xml");
			var values = doc.getElementsByTagName("t:MimeContent");
			//var subject = doc.getElementsByTagName("t:Subject");
			//console.log(values[0].textContent);
			
			
			var formData = new FormData();

			formData.append("date", currentDate);

			formData.append("subject", subject); 

			formData.append("attachEmail", base64toBlob(values[0].textContent));

			var request = new XMLHttpRequest();
			request.open("POST", "https://informatik.hs-bremerhaven.de/ykapuler/Praxis/add_in_outlook.php");
			request.send(formData);			
			
			//Office.context.ui.closeContainer();
			
		});
		setTimeout(() => {
			var mapForm = document.createElement("form");
			mapForm.target = "_blank";    
			mapForm.method = "POST";
			mapForm.action = "https://informatik.hs-bremerhaven.de/ykapuler/Praxis/check_DB.php";
	
			var mapInput = document.createElement("input");
			mapInput.type = "text";
			mapInput.name = "date";
			mapInput.value = currentDate;
			
			/*var mapInput2 = document.createElement("input");
			mapInput2.type = "text";
			mapInput2.name = "subject";
			mapInput2.value = subject;*/
	
			mapForm.appendChild(mapInput);
			//mapForm.appendChild(mapInput2);

			document.body.appendChild(mapForm);

			mapForm.submit();
			
			
			
		}, 1500);
		
		
	};

	function getActualDate(){
		var aDate = new Date();
		
		var month, day, hours, minutes, seconds;
		
		
		month = ((aDate.getMonth()+1) < 10)? "0" + (aDate.getMonth()+1): aDate.getMonth()+1;
		
		day = (aDate.getDate() < 10)? "0" + aDate.getDate(): aDate.getDate();
		
		hours = (aDate.getHours() < 10)? "0" + aDate.getHours(): aDate.getHours();
		
		minutes = (aDate.getMinutes() < 10)? "0" + aDate.getMinutes(): aDate.getMinutes();
		
		seconds = (aDate.getSeconds() < 10)? "0" + aDate.getSeconds(): aDate.getSeconds();
		
		
		var actualDate = aDate.getFullYear() + "-" + month + "-" + day + " "
		+ hours + ":" + minutes + ":" + seconds;
		
		return actualDate;
	}

	function GetItem() {
		var results =
			'  <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
			'    <ItemShape>' +
			'      <t:BaseShape>IdOnly</t:BaseShape>' +
			'      <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
			'      <AdditionalProperties xmlns="http://schemas.microsoft.com/exchange/services/2006/types">' +
			'        <FieldURI FieldURI="item:Subject" />' +
			'      </AdditionalProperties>' +
			'    </ItemShape>' +
			'    <ItemIds>' +
			'      <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" />' +
			'    </ItemIds>' +
			'  </GetItem>';
 
		return results;
	}
	
	function getSoapEnvelope(request) {
    // Wrap an Exchange Web Services request in a SOAP envelope.
		var result =

		'<?xml version="1.0" encoding="utf-8"?>' +
		'<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
		'               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
		'               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
		'               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
		'  <soap:Header>' +
		'    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
		'  </soap:Header>' +
		'  <soap:Body>' +

		request +

		'  </soap:Body>' +
		'</soap:Envelope>';
		
		return result;
	}
	
	function base64toBlob(base64Data, contentType) {
		contentType = contentType || '';
		var sliceSize = 1024;
		var byteCharacters = atob(base64Data);
		var bytesLength = byteCharacters.length;
		var slicesCount = Math.ceil(bytesLength / sliceSize);
		var byteArrays = new Array(slicesCount);

		for (var sliceIndex = 0; sliceIndex < slicesCount; ++sliceIndex) {
			var begin = sliceIndex * sliceSize;
			var end = Math.min(begin + sliceSize, bytesLength);
	
			var bytes = new Array(end - begin);
			for (var offset = begin, i = 0; offset < end; ++i, ++offset) {
				bytes[i] = byteCharacters[offset].charCodeAt(0);
			}
			byteArrays[sliceIndex] = new Uint8Array(bytes);
		}
		return new Blob(byteArrays, { type: contentType });
	}
	
})();
