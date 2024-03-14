'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            loadItemProps(Office.context.mailbox.item);
        });
    });
    function sendEmailDataToAPI(formData) {
        console.log("Inside Send Email");
        for (let pair of formData.entries()) {
            // pair[0] contains the key, pair[1] contains the value
            console.log(pair[0] + ', ' + pair[1]);
            // Make HTTP request to API
        }
        fetch("https://helpdesk.hindujatech.com:3695/ticketing_tool/add_ticket_by_mail", {
            method: "POST",
            /*  headers: {
                  "Content-Type": "application/json"
              },*/
            body: formData //JSON.stringify(emailData)
        })
            .then(response => {
                if (response.ok) {
                    return response.json();
                } else {
                    throw new Error("Failed to send email data to API");
                }
            })
            .then(data => {
                console.log("API response:", data);
                // Handle success response from API
                var message = data.message || ""; // Ensure message is always a string
                if (message.trim() !== '') { // Check if message is not empty
                    // Display the message in a dialog
                    $('<div></div>').html('<p>' + message + '</p>').dialog({
                        title: 'Alert',
                        modal: true,
                        width: 'auto',
                        open: function (event, ui) {
                        },
                        buttons: [{
                            text: "Close",
                            click: function () {
                                $(this).dialog('close');
                            },
                            class: "custom-close-button"
                        }]

                    });
                }

            })
            .catch(error => {
                console.error("Error:", error);
                // Handle error
            });
    }


    function loadItemProps(item) {
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-mailto').text(item.to);
        $('#item-date').text(item.dateTimeCreated);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        console.log(`Creation date and time: ${Office.context.mailbox.item.dateTimeCreated}`);

        var toRecipients = Office.context.mailbox.item.to;

        // Clear previous content
        $('#item-mailto').empty();

        // Iterate over 'to' recipients and set HTML content
        toRecipients.forEach(function (recipient) {
            $('#item-mailto').append(`<p>${recipient.displayName} &lt;${recipient.emailAddress}&gt;</p>`);
        });

        // Get attachments
        var attachments = Office.context.mailbox.item.attachments;

        // Log attachment details
        attachments.forEach(function (attachment) {
            console.log('Attachment:', attachment.name);
            $('#item-attachment').append(`<p>${attachment.name}</p>`);
        });

        var attachments = Office.context.mailbox.item.attachments;
        if (attachments.length > 0) {
            var attachment = attachments[0];
            //  attachments.forEach(function (attachment) {
            var attachmentName = attachment.name;
            var attachmentId = attachment.id;
            var fileExtension = attachmentName.split('.').pop().toLowerCase();
            // Now you have the file extension in 'fileExtension'
            console.log("File extension:", fileExtension);
            if (fileExtension != '.gif') {
                var restUrl = Office.context.mailbox.restUrl + '/v2.0/me/messages/' + Office.context.mailbox.item.itemId +
                    '/attachments/' + attachmentId + '/$value';
                console.log("restUrl " + restUrl);

                // Get the callback token
                Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
                    if (result.error) {
                        console.error("Failed to get callback token:", result.error.message);
                        return;
                    }

                    var accessToken = result.value;

                    // Make a request to download the attachment content
                    fetch(restUrl, {
                        method: "GET",
                        headers: {
                            "Authorization": "Bearer " + accessToken
                        }
                    })
                        .then(response => {
                            if (response.ok) {
                                return response.blob(); // Get the attachment content as a Blob
                            } else {
                                var formData = new FormData();
                                formData.append('internetMessageId', item.internetMessageId);
                                //  formData.append('attachment', attachmentFile);
                                formData.append('from', item.from.emailAddress);
                                formData.append('subject', item.subject);

                                var toRecipientsJSON = JSON.stringify(toRecipients.map(recipient => ({
                                    "emailAddress": recipient.emailAddress
                                })));
                                // Append the 'to' recipients JSON string to the FormData object
                                formData.append('mail-to', toRecipientsJSON);
                                Office.context.mailbox.item.body.getAsync("text", function (result) {
                                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                                        console.log("Body " + result.value);
                                        formData.append('body', JSON.stringify(result.value));
                                        $('#item-body').text(result.value);
                                        sendEmailDataToAPI(formData);

                                    } else {
                                        console.error("Failed to get email body:", result.error);
                                    }
                                });

                                throw new Error("Failed to download attachment content.");
                            }
                        })
                        .then(blob => {
                            // emailData.fileContent = blob;

                            // Handle the attachment content
                            // Depending on the file type, you may need different handling (e.g., display, save, process)
                            // For example, if it's a PDF, you might display it in an iframe or download it
                            // For now, let's just log the Blob object
                            var reader = new FileReader();
                            reader.readAsDataURL(blob);

                            reader.onloadend = function () {
                                var base64Data = reader.result.split(',')[1];

                                var attachmentFile = new File([blob], attachment.name);

                                // Create FormData object
                                var formData = new FormData();
                                formData.append('internetMessageId', item.internetMessageId);
                                formData.append('attachment', attachmentFile);
                                formData.append('from', item.from.emailAddress);
                                formData.append('subject', item.subject);

                                var toRecipientsJSON = JSON.stringify(toRecipients.map(recipient => ({
                                    "emailAddress": recipient.emailAddress
                                })));

                                // Append the 'to' recipients JSON string to the FormData object
                                formData.append('mail-to', toRecipientsJSON);

                                Office.context.mailbox.item.body.getAsync("text", function (result) {
                                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                                        console.log("Body " + result.value);
                                        formData.append('body', JSON.stringify(result.value));
                                        $('#item-body').text(result.value);
                                        sendEmailDataToAPI(formData);

                                    } else {
                                        console.error("Failed to get email body:", result.error);
                                    }
                                });

                            };
                        })
                        .catch(error => {
                            console.error("Error:", error);
                        });
                });


            } else if (fileExtension === '.gif') {

                var formData = new FormData();
                formData.append('internetMessageId', item.internetMessageId);
                //  formData.append('attachment', attachmentFile);
                formData.append('from', item.from.emailAddress);
                formData.append('subject', item.subject);

                var toRecipientsJSON = JSON.stringify(toRecipients.map(recipient => ({
                    "emailAddress": recipient.emailAddress
                })));
                // Append the 'to' recipients JSON string to the FormData object
                formData.append('mail-to', toRecipientsJSON);
                Office.context.mailbox.item.body.getAsync("text", function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Body " + result.value);
                        formData.append('body', JSON.stringify(result.value));
                        $('#item-body').text(result.value);
                        sendEmailDataToAPI(formData);

                    } else {
                        console.error("Failed to get email body:", result.error);
                    }
                });

            }    // End of if xlsx
        }

    }
}// End of function
)();