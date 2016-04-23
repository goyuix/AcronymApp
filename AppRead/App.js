/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};
    app.acronyms = {};
    var ACRONYMS = 'acronyms';

    // get data from REST API to lookup acronyms
    app.loadData = function (url) {
        jQuery.ajax({
            url: url,
            headers: { 'accept': 'application/json;odata=verbose' },
            success: function (response) {
                // process results
                var item = null;
                if (response && response.d && response.d.results) {
                    for (var i = 0; i < response.d.results.length; i++) {
                        item = response.d.results[i];
                        app.acronyms[item.Acronym] = item.Title;
                    }
                }
                // if there is more paged data, start request for it
                if (response && response.d && response.d["__next"]) {
                    app.loadData(response.d["__next"]);
                } else {
                    window.sessionStorage.setItem(ACRONYMS, JSON.stringify(app.acronyms));
                    app.showNotification("Completed loading " + Object.keys(app.acronyms).length + " acronyms", "");
                }
            }
        });
    };

    // Common initialization function (to be called from each page)
    app.initialize = function () {
        // begin loading acronym data
        app.acronyms = window.sessionStorage.getItem(ACRONYMS);
        if (!app.acronyms || !app.acronyms.length || app.acronyms.length < 1) {
            app.acronyms = {};
            app.loadData("https://www.wecc.biz/_api/Web/Lists/getByTitle('Acronyms')/items?$select=Acronym,Title");
        }

        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();
