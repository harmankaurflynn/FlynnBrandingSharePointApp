/*
==========================
Vertical Responsive Menu
==========================
*/

'use strict';
var ns = CreateNamespace('FlynnBranding');



ns.LeftNav = function () { };
// How long shall we cache the HTML response? (specify a Prime Number)

ns.LeftNav.ExpirationTimeoutInMinutes = "13";



$(window).load(function () {
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

                            //$("#sideNavBox").append("<header class='header clearfix'><button type='button' id='toggleMenu'class='toggle_menu'><iclass='fa fa-bars'></i></button></header><nav class='vertical_nav'><ul id='js-menu' class='menu'><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Safety Policies'><i class='menu--icon  fa fa-fw fa-user'></i><span class='menu--label'>Safety Policies</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='http://google.com' target='_blank' class='sub_menu--link sub_menu--link__active'>Submenu</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Submenu</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Submenu</a></li></ul></li><li class='menu--item'><a href='#' class='menu--link' title='Health & Safety Management'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Health & Safety Management</span></a></li><li class='menu--item'><a href='#' class='menu--link' title='Safe Work Procedures'><i class='menu--icon  fa fa-fw fa-cog'></i><span class='menu--label'>Safe Work Procedures</span></a></li><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Corporate'><i class='menu--icon  fa fa-fw fa-database'></i><span class='menu--label'>Corporate</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Executive Safety Team</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Communications</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Health & Safety Department</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Safety Recognition</a></li></ul></li><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Region/Branch'><i class='menu--icon  fa fa-fw fa-database'></i><span class='menu--label'>Region/Branch</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='#' class='sub_menu--link'>HS Plan</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>HS Annual Goals</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Safety Reviews</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Reporting(KPI)</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>JHS Committees</a></li></ul></li><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Region/Branch'><i class='menu--icon  fa fa-fw fa-database'></i><span class='menu--label'>Site/Project</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Site HS Plan</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Hazard Analysis</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Flynn Stores</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Competency</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Subcontractor Management</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Incident/Case Management</a></li></ul></li><li class='menu--item'><a href='#' class='menu--link' title='Care'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Care</span></a></li><li class='menu--item'><a href='#' class='menu--link' title='Safety Day'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Safety Day</span></a></li><li class='menu--item'><a href='#' class='menu--link' title='Flynn Family'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Flynn Family</span></a></li></ul><button id='collapse_menu' class='collapse_menu'><i class='collapse_menu--icon  fa fa-fw'></i><span class='collapse_menu--label'>Recolher menu</span></button></nav>");


                            

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

ns.LeftNav.GetContents = function ()

{
    

    // Save the current control content as the fallback content in case we encounter any errors during processing; we will simply re-render the current content.

    var currentHtml = $('#sideNavBox').html();

    var fallbackHtml = currentHtml;



    // If current control content is not present, use default Html; optionally, insert a progress indicator while we request the BDO for the control. 

    if (currentHtml == "")

    {
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



    try

    {
       
        // Request the BDO for the control. We use DurableStorage for Global Nav -- its content is not personalized/private

        ns.BusinessDataAccess.GetLeftNavMenuData({ storageMode: ns.StorageManager.DurableStorageMode, useSlidingExpiration: false, timeout: ns.LeftNav.ExpirationTimeoutInMinutes }).then(



            function (leftNavData)

            {
               

                //  construct the content HTML for the control using the data returned in the BDO

                var contentHtml = "<header class=\"header clearfix\"><button type=\"button\" id=\"toggleMenu\" class=\"toggle_menu\"><iclass=\"fa fa-bars\"></i></button></header>"+
                    "<nav class=\"vertical_nav\">"+
                    "<ul id=\"js-menu\" class=\"menu\">";
                   
               



                if (leftNavData == null || leftNavData.Type == ns.BusinessDataAccess.ErrorDataType)

                {
                    //alert("leftnavdata null")

                    // For some reason, a valid BDO was not returned; update the control instance with the fallback Html.

                    contentHtml = fallbackHtml;

                }

                else

                {
                   // alert("left nav not null");
                    // We have received a valid BDO; update the control instance with the BDO data.

                    if (leftNavData.Nodes.length > 0)

                    {
                        var leftNavMenu = leftNavData.Nodes.filter(function (p) { return p.Parent === undefined; });
                       // var leftNavMenu = leftNavData.Nodes.filter(p => (p.Parent ===undefined));
                       

                        var listItemInfo = "";

                        $.each(leftNavMenu, function ()

                        {

                            var linkText = this.Title;

                            var linkUrl = this.Url;
                            var blHasSubMenu = false;
                            //var leftNavSubMenu = leftNavData.Nodes.filter(p => (p.Parent == this.Title));
                            var leftNavSubMenu = leftNavData.Nodes.filter(function (p) { return p.Parent === linkText; });
                            if (leftNavSubMenu && leftNavSubMenu.length > 0)
                                blHasSubMenu = true;

                            if (blHasSubMenu)
                                listItemInfo +=
                              "<li class=\"menu--item  menu--item__has_sub_menu\">"
                            else
                                listItemInfo +=
                              "<li class=\"menu--item\">"


                            listItemInfo +=
                                "<label class=\"menu--link\" title=\"" + linkText + "\"><i class=\"menu--icon  fa fa-fw fa-user\"></i><span class=\"menu--label\">" + linkText + "</span></label>"

                           

                            if (blHasSubMenu) {
                               

                                listItemInfo +=
                                "<ul class=\"sub_menu\">"

                                $.each(leftNavSubMenu, function () {

                                    var linkText = this.Title;

                                    var linkUrl = this.Url;

                                    listItemInfo +=
                                        "<li class=\"sub_menu--item\"><a href=\"" + linkUrl + "\" target=\"_blank\" class=\"sub_menu--link sub_menu--link__active\">" + linkText + "</a></li>"

                                });

                                listItemInfo += "</ul>"
                            }

                            listItemInfo += "</li>"

                           

                          

                        });

                       

                        contentHtml += listItemInfo.toString();

                        contentHtml += "</ul><button id=\"collapse_menu\" class=\"collapse_menu\">"+
                        "<i class=\"collapse_menu--icon  fa fa-fw\"></i><span class=\"collapse_menu--label\">Recolher menu</span></button></nav>";
                        
                    }

                    else

                    {

                        // We have received an empty BDO; update the control instance with the default HTML.

                        contentHtml = ns.LeftNav.DefaultHtml

                    }

                }



                // Update the control instance with the resulting content HTML

                $('#sideNavBox').empty();

                $('#sideNavBox').append(contentHtml);
              
                applyCSS();
               
            },



            function (sender, args)

            {

                ns.LogError(args.get_message());



                // The BDO request failed; update the control instance with the fallback Html.

                $('#sideNavBox').empty();

                $('#sideNavBox').append(fallbackHtml);

            }

        );

    }

    catch (ex)

    {

        // The BDO request failed; update the control instance with the fallback Html.

        $('#sideNavBox').empty();

        $('#sideNavBox').append(fallbackHtml);

    }

}




var applyCSS=(function () {
   
    if (document.readyState !== 'complete') {
        alert(document.readyState);
        
        applyCSS();
    }
    else {
        // clearInterval( tid );

        $("head").append("<link rel='stylesheet' href='https://dspdev05webapp.flynncompanies.com/sites/AppModelTest/Style%20Library/font-awesome.css' type='text/css' >");
        $("head").append("<link rel='stylesheet' href='https://dspdev05webapp.flynncompanies.com/sites/AppModelTest/Style%20Library/vertical-responsive-menu.css' type='text/css' >");
        $("head").append("<link rel='stylesheet' href='https://dspdev05webapp.flynncompanies.com/sites/AppModelTest/Style%20Library/normalize.css' type='text/css' >");
        var querySelector = document.querySelector.bind(document);

        var nav = document.querySelector('.vertical_nav');
        var wrapper = document.querySelector('.wrapper');

        var menu = document.getElementById("js-menu");
        

        var subnavs = menu.querySelectorAll('.menu--item__has_sub_menu');

        // Toggle menu click
        querySelector('.toggle_menu').onclick = function () {

            nav.classList.toggle('vertical_nav__opened');

            wrapper.classList.toggle('toggle-content');

        };


        // Minify menu on menu_minifier click
        querySelector('.collapse_menu').onclick = function () {

            nav.classList.toggle('vertical_nav__minify');

            wrapper.classList.toggle('wrapper__minify');

            for (var j = 0; j < subnavs.length; j++) {
                subnavs[j].classList.remove('menu--subitens__opened');
            }

        };


        // Open Sub Menu

        for (var i = 0; i < subnavs.length; i++) {

            if (subnavs[i].classList.contains('menu--item__has_sub_menu')) {

                subnavs[i].querySelector('.menu--link').addEventListener('click', function (e) {

                    for (var j = 0; j < subnavs.length; j++) {

                        if (e.target.offsetParent != subnavs[j])
                            subnavs[j].classList.remove('menu--subitens__opened');

                    }

                    e.target.offsetParent.classList.toggle('menu--subitens__opened');

                }, false);

            }
        }
    }


});

