/* Common app functionality */

const app = ( () => {
    const app = {};
    app.bindingID = 'myBinding';

    // Common initialization function (to be called from each page).
    app.initialize = () => {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').on('click', () => {{
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function.
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };
    };

    return app;
})();