/*
==========================
Vertical Responsive Menu
==========================
*/

'use strict';
var ns = CreateNamespace('FlynnBranding');



ns.LeftNav = function () { };
// How long shall we cache the HTML response? (specify a Prime Number)

ns.LeftNav.ExpirationTimeoutInMinutes = "0";


console.log("start " + new Date().toLocaleString());
$(document).ready(function () {
    console.log("load " + new Date().toLocaleString());
    $.getScript("/_layouts/15/MicrosoftAjax.js", function () {
        $.getScript("/_layouts/15/SP.Runtime.js", function () {
            $.getScript("/_layouts/15/SP.js", function () {


                //Type.registerNamespace('FlynnBranding');
                // FlynnBranding.LeftNavigation = FlynnBranding.LeftNavigation || {};

                ns.LeftNav.Render = function () {

                    var webProperties,

                    getWebPropertiesSucceeded = function () {

                        {

                            ns.LeftNav.GetContents();

                            // $("#sideNavBox").empty();



                        }



                    },

                    handleError = function () {
                        alert('Error rendering header/footer');
                    }

                    return {
                        initializeLeftNav: function () {

                            var clientContext = SP.ClientContext.get_current();
                            var hostWeb = clientContext.get_web();
                            webProperties = hostWeb.get_allProperties();
                            clientContext.load(webProperties);
                            clientContext.executeQueryAsync(getWebPropertiesSucceeded, handleError);
                        }
                    }
                }();

                if (typeof _spPageContextInfo != "undefined" && _spPageContextInfo != null) {
                    // MDS enabled

                    RegisterModuleInit(_spPageContextInfo.siteServerRelativeUrl + 'Style Library/leftNavMenu.js', ns.LeftNav.Render.initializeHeaderFooter);
                }
                // Run now on this page (and non-MDS scenarios)
                //ExecuteOrDelayUntilScriptLoaded(Flynn.LeftNavigation.Render.initializeLeftNav, "SP.Runtime.js"); 

                ns.LeftNav.Render.initializeLeftNav();

            });
        });
    });
});

// Constructs the entire HTML rendering of the Control and injects the HTML into the control

ns.LeftNav.GetContents = function () {


    // Save the current control content as the fallback content in case we encounter any errors during processing; we will simply re-render the current content.

    var currentHtml = $('#sideNavBox').html();

    var fallbackHtml = currentHtml;

    $('#DeltaPlaceHolderMain').addClass('wrapper');



    // If current control content is not present, use default Html; optionally, insert a progress indicator while we request the BDO for the control. 

    if (currentHtml == "") {
        alert("empty");
        // Since current control content is not present, use the Default Html as the fallback content.

        fallbackHtml = ns.GlobalNav.DefaultHtml;



        ////----------------------------------------------------------------------------

        //// Note: ENABLE this code block if you wish to present a progress indicator...

        ////----------------------------------------------------------------------------

        //// insert a progress indicator while we request the BDO for the control. 

        //$('#pnpGlobalNav').empty();

        //$('#pnpGlobalNav').append(

        //  "<ul class=\"ms-core-suiteLinkList\">" + 

        //    "<li class=\"ms-core-suiteLink\">" +

        //      "<a class=\"ms-core-suiteLink-a\" href=\"/\"><img src=\"" + ns.Configuration.PortalCdnUrl + "/images/loading.gif\" alt=\"loading...\"/>&nbsp;Working on it...</a>"); +

        //    "</li>" +

        //  "</ul>";

        //----------------------------------------------------------------------------

    }



    try {

        // Request the BDO for the control. We use DurableStorage for Global Nav -- its content is not personalized/private

        ns.BusinessDataAccess.GetLeftNavMenuData({ storageMode: ns.StorageManager.DurableStorageMode, useSlidingExpiration: false, timeout: ns.LeftNav.ExpirationTimeoutInMinutes }).then(




            function (leftNavData) {

                var currentUrl = window.location.href;
                if (currentUrl.indexOf('/') > 0)
                    currentUrl = currentUrl.substring(currentUrl.indexOf('/') + 1);

                //  construct the content HTML for the control using the data returned in the BDO

                var contentHtml = "<header class=\"header clearfix\"><button type=\"button\" id=\"toggleMenu\" class=\"toggle_menu\"><i class=\"fa fa-bars\"></i></button></header>" +
                    "<nav class=\"vertical_nav\">" +
                    "<ul id=\"js-menu\" class=\"menu\">";





                if (leftNavData == null || leftNavData.Type == ns.BusinessDataAccess.ErrorDataType) {
                    //alert("leftnavdata null")

                    // For some reason, a valid BDO was not returned; update the control instance with the fallback Html.

                    contentHtml = fallbackHtml;

                }

                else {
                    // alert("left nav not null");
                    // We have received a valid BDO; update the control instance with the BDO data.

                    var nodeIndex = 0;

                    if (leftNavData.Nodes.length > 0) {
                        var leftNavMenu = leftNavData.Nodes.filter(function (p) { return p.Parent === undefined; });
                        // var leftNavMenu = leftNavData.Nodes.filter(p => (p.Parent ===undefined));


                        var listItemInfo = "";

                        $.each(leftNavMenu, function () {

                            var linkText = this.Title;

                            var linkUrl = "#";

                            if (this.Url)
                                linkUrl = this.Url.Url;


                            var blHasSubMenu = false;
                            //var leftNavSubMenu = leftNavData.Nodes.filter(p => (p.Parent == this.Title));
                            var leftNavSubMenu = leftNavData.Nodes.filter(function (p) { return p.Parent === linkText; });
                            if (leftNavSubMenu && leftNavSubMenu.length > 0)
                                blHasSubMenu = true;

                            if (blHasSubMenu)
                                listItemInfo +=
                              "<li class=\"menu--item  menu--item__has_sub_menu  menu--subitens__opened\">"
                            else

                                listItemInfo +=
                              "<li class=\"menu--item\">"


                            if (nodeIndex == 0)
                                listItemInfo +=
                               "<label class=\"menu--label menu--label__home\" title=\"" + linkText + "\"><i class=\"menu--icon  fa fa-fw fa-user\"></i>"
                            else
                                listItemInfo +=
                                    "<label class=\"menu--label\" title=\"" + linkText + "\"><i class=\"menu--icon  fa fa-fw fa-user\"></i>"

                            if (linkUrl == window.location.href)
                                listItemInfo += "<a class=\"menu--link menu--link__active\" href=\"" + linkUrl + "\">" + linkText + "</a></label>";
                            else

                                listItemInfo += "<a class=\"menu--link\" href=\"" + linkUrl + "\">" + linkText + "</a></label>";

                            if (blHasSubMenu) {


                                listItemInfo +=
                                "<ul class=\"sub_menu\">"

                                $.each(leftNavSubMenu, function () {

                                    var linkText = this.Title;

                                    var linkUrl = "#";

                                    if (this.Url)
                                        linkUrl = this.Url.Url;


                                    //if (linkUrl == window.location.href)
                                    // listItemInfo += "<li class=\"sub_menu--item\"><a href=\"" + linkUrl + "\" target=\"_blank\" class=\"sub_menu--link sub_menu--link__active\">" + linkText + "</a></li>";
                                    // else

                                    // listItemInfo += "<li class=\"sub_menu--item\"><a href=\"" + linkUrl + "\" target=\"_blank\" class=\"sub_menu--link\">" + linkText + "</a></li>";

                                    var blHasSubSubMenu = false;
                                    var leftNavSubSubMenu = leftNavData.Nodes.filter(function (p) { return p.Parent === linkText; });
                                    if (leftNavSubSubMenu && leftNavSubSubMenu.length > 0)
                                        blHasSubSubMenu = true;

                                    if (blHasSubSubMenu)
                                        listItemInfo +=
                                        "<li class=\"sub_menu--item sub_menu--item__has_sub_menu\">"
                                    else

                                        listItemInfo +=
                                      "<li class=\"sub_menu--item\">"
                                    listItemInfo +=
                                "<label class=\"menu--label\" title=\"" + linkText + "\">"
                                    if (linkUrl == window.location.href)
                                        listItemInfo += "<a href=\"" + linkUrl + "\"  class=\"sub_menu--link sub_menu--link__active\">" + linkText + "</a></label>";
                                    else
                                        listItemInfo += "<a href=\"" + linkUrl + "\"  class=\"sub_menu--link\">" + linkText + "</a></label>";

                                    if (blHasSubSubMenu) {


                                        listItemInfo +=
                                        "<ul class=\"sub_sub_menu\">"

                                        $.each(leftNavSubSubMenu, function () {
                                            var linkText = this.Title;

                                            var linkUrl = "#";

                                            if (this.Url)
                                                linkUrl = this.Url.Url;
                                            if (linkUrl == window.location.href)
                                                listItemInfo += "<li class=\"sub_sub_menu--item\"><a href=\"" + linkUrl + "\" target=\"_blank\" class=\"sub_sub_menu--link sub_sub_menu--link__active\">" + linkText + "</a></li>";
                                            else
                                                listItemInfo += "<li class=\"sub_sub_menu--item\"><a href=\"" + linkUrl + "\" target=\"_blank\" class=\"sub_sub_menu--link\">" + linkText + "</a></li>";
                                        });
                                        listItemInfo += "</ul>"
                                    }

                                    listItemInfo += "</li>"
                                });

                                listItemInfo += "</ul>"
                            }

                            listItemInfo += "</li>"


                            nodeIndex++;


                        });



                        contentHtml += listItemInfo.toString();

                        contentHtml += "</ul></nav>";

                    }

                    else {

                        // We have received an empty BDO; update the control instance with the default HTML.

                        contentHtml = ns.LeftNav.DefaultHtml

                    }

                }



                // Update the control instance with the resulting content HTML

                $('#sideNavBox').empty();

                $('#sideNavBox').append(contentHtml);

                applyCSS();

            },



            function (sender, args) {

                ns.LogError(args.get_message());



                // The BDO request failed; update the control instance with the fallback Html.

                $('#sideNavBox').empty();

                $('#sideNavBox').append(fallbackHtml);

            }

        );

    }

    catch (ex) {

        // The BDO request failed; update the control instance with the fallback Html.

        $('#sideNavBox').empty();

        $('#sideNavBox').append(fallbackHtml);

    }

}




var applyCSS = (function () {


    if (document.readyState !== 'complete') {
        alert(document.readyState);

        applyCSS();
    }
    else {
        // clearInterval( tid );
        var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl;

        $("head").append("<link rel='stylesheet' href='" + siteURL + "/Style%20Library/FlynnBranding/CSS/font-awesome.css' type='text/css' >");
        $("head").append("<link rel='stylesheet' href='" + siteURL + "/Style%20Library/FlynnBranding/CSS/vertical-responsive-menu.css' type='text/css' >");
        $("head").append("<link rel='stylesheet' href='" + siteURL + "/Style%20Library/FlynnBranding/CSS/normalize.css' type='text/css' >");
        $("head").append("<link rel='stylesheet' href='" + siteURL + "/Style%20Library/FlynnBranding/CSS/flynnbranding.css' type='text/css' >");
        var querySelector = document.querySelector.bind(document);

        var nav = document.querySelector('.vertical_nav');
        var wrapper = document.querySelector('.wrapper');

        var menu = document.getElementById("js-menu");


        var subnavs = menu.querySelectorAll('.menu--item__has_sub_menu');
        var subsubnavs = menu.querySelectorAll('.sub_menu--item__has_sub_menu');

        // Toggle menu click
        querySelector('.toggle_menu').onclick = function () {

            nav.classList.toggle('vertical_nav__opened');

            wrapper.classList.toggle('toggle-content');

        };


        // Minify menu on menu_minifier click
        //querySelector('.collapse_menu').onclick = function () {

        //    nav.classList.toggle('vertical_nav__minify');

        //    wrapper.classList.toggle('wrapper__minify');

        //    for (var j = 0; j < subnavs.length; j++) {
        //        //subnavs[j].classList.remove('menu--subitens__opened');
        //    }


        //    for (var j = 0; j < subsubnavs.length; j++) {
        //       // subsubnavs[j].classList.remove('sub_menu--subitens__opened');
        //    }

        //};


        // Open Sub Menu

        for (var i = 0; i < subnavs.length; i++) {

            if (subnavs[i].classList.contains('menu--item__has_sub_menu')) {

                subnavs[i].querySelector('.menu--link').addEventListener('click', function (e) {

                    for (var j = 0; j < subnavs.length; j++) {

                        if (e.target.offsetParent != subnavs[j]) {
                            // subnavs[j].classList.remove('menu--subitens__opened');
                        }

                    }

                    e.target.offsetParent.classList.toggle('menu--subitens__opened');

                }, false);

                for (var j = 0; j < subnavs.length; j++) {
                    // subnavs[j].classList.toggle('menu--subitens__opened');

                }


                var selectedLink = $('.sub_menu--link__active');


            }
        }

        for (var i = 0; i < subsubnavs.length; i++) {

            if (subsubnavs[i].classList.contains('sub_menu--item__has_sub_menu')) {



                subsubnavs[i].querySelector('.sub_menu--link').addEventListener('click', function (e) {

                    for (var j = 0; j < subsubnavs.length; j++) {

                        if (e.target.offsetParent != subsubnavs[j])
                            subsubnavs[j].classList.remove('sub_menu--subitens__opened');


                    }

                    e.target.offsetParent.classList.toggle('sub_menu--subitens__opened');

                }, false);

                for (var j = 0; j < subsubnavs.length; j++) {
                    // subnavs[j].classList.toggle('menu--subitens__opened');

                }


                var selectedLink = $('.sub_menu--link__active');


            }
        }

        var selectedLink = $('.sub_menu--link__active');
        if (selectedLink != null && selectedLink.length > 0)
            selectedLink[0].parentNode.parentNode.parentNode.classList.toggle('menu--subitens__opened');

        var selectedsubLink = $('.sub_sub_menu--link__active');
        if (selectedsubLink != null && selectedsubLink.length > 0) {
            selectedsubLink[0].parentNode.parentNode.parentNode.classList.toggle('sub_menu--subitens__opened');
            selectedsubLink[0].parentNode.parentNode.parentNode.parentNode.parentNode.classList.toggle('menu--subitens__opened');
        }

    }


});