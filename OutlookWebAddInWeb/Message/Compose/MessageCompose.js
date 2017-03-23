(function () {
    "use strict";

    var item;
    var subject;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            item = Office.context.mailbox.item;

            getResultsForSubject();            

            //Bind Search
            $("#searchform").submit(function (event) {
                $('#mainnav').removeClass('in');
                $("#queryresults").empty();
                var searchFor = encodeURIComponent($('#query').val());
                getResults(searchFor);
                event.preventDefault();
            });

        });
    };

    function getResultsForSubject() {
        item.subject.getAsync(
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed)
                {
                    console.log(asyncResult.error.message);
                }
                else
                {
                    var subjectText = asyncResult.value;
                    subjectText = encodeURIComponent(subjectText.replace("Re:", ""));
                    getResults(subjectText);                    
                }
            });
    }

    function getResults(query) {

        var myUrl = "/bingsearch/get/" + query;

        $.ajax({
            url: myUrl,
            type: "GET"
            })
            .done(function (data) {
                var len = data.length
                for (var i = 0; i < len; i++)
                {
                    $("#queryresults").append('<a href="' + data[i].Url + '" target="_blank" class="list-group-item"><h5 class="list-group-item-heading">' + data[i].Name + '</h5><p class="list-group-item-text">' + data[i].Snippet + '</p></a>');
                }
                $('#numresults').html(len);

            })
            .fail(function (err) {
                console("error " + err);
            });
    }

})();