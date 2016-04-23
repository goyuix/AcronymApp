﻿/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            displayItemDetails();
        });
    };

    // Displays the "Subject" and "From" fields, based on the current mail item
    function displayItemDetails() {
        var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        $('#subject').text(item.subject);

        var from;
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            from = Office.cast.item.toMessageRead(item).from;
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            from = Office.cast.item.toAppointmentRead(item).organizer;
        }

        if (from) {
            $('#from').text(from.displayName);
            $('#from').click(function () {
                app.showNotification(from.displayName, from.emailAddress);
            });
        }

        Office.context.mailbox.item.body.getAsync("text", {}, function (result) {
            var html = [];
            var matches = $.unique(result.value.match(/\b[A-Z]{3,}\b/g));
            for (var i=0;i<matches.length;i++) {
                html.push('<li><b>'+matches[i]+'</b><br/>'+(app.acronyms[matches[i]] ? app.acronyms[matches[i]] : 'No matching definition'))+'</li>';
            }
            $("#body").text('<ul>'+html.join(',')+'</ul>');
        })
        
    }
})();
