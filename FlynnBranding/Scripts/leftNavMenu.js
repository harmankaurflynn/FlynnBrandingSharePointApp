/*
==========================
Vertical Responsive Menu
==========================
*/

'use strict';

$(window).load(function () {
    $.getScript("/_layouts/15/MicrosoftAjax.js", function () {
        $.getScript("/_layouts/15/SP.Runtime.js", function () {
            $.getScript("/_layouts/15/SP.js", function () {
                Type.registerNamespace('Flynn');
                Flynn.LeftNavigation = Flynn.LeftNavigation || {};

                Flynn.LeftNavigation.Render = function () {

                    var webProperties,

                    getWebPropertiesSucceeded = function () {

                        {

                            $("#sideNavBox").empty();

                            $("#sideNavBox").append("<header class='header clearfix'><button type='button' id='toggleMenu'class='toggle_menu'><iclass='fa fa-bars'></i></button></header><nav class='vertical_nav'><ul id='js-menu' class='menu'><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Safety Policies'><i class='menu--icon  fa fa-fw fa-user'></i><span class='menu--label'>Safety Policies</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='http://google.com' target='_blank' class='sub_menu--link sub_menu--link__active'>Submenu</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Submenu</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Submenu</a></li></ul></li><li class='menu--item'><a href='#' class='menu--link' title='Health & Safety Management'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Health & Safety Management</span></a></li><li class='menu--item'><a href='#' class='menu--link' title='Safe Work Procedures'><i class='menu--icon  fa fa-fw fa-cog'></i><span class='menu--label'>Safe Work Procedures</span></a></li><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Corporate'><i class='menu--icon  fa fa-fw fa-database'></i><span class='menu--label'>Corporate</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Executive Safety Team</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Communications</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Health & Safety Department</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Safety Recognition</a></li></ul></li><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Region/Branch'><i class='menu--icon  fa fa-fw fa-database'></i><span class='menu--label'>Region/Branch</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='#' class='sub_menu--link'>HS Plan</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>HS Annual Goals</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Safety Reviews</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Reporting(KPI)</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>JHS Committees</a></li></ul></li><li class='menu--item  menu--item__has_sub_menu'><label class='menu--link' title='Region/Branch'><i class='menu--icon  fa fa-fw fa-database'></i><span class='menu--label'>Site/Project</span></label><ul class='sub_menu'><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Site HS Plan</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Hazard Analysis</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Flynn Stores</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Competency</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Subcontractor Management</a></li><li class='sub_menu--item'><a href='#' class='sub_menu--link'>Incident/Case Management</a></li></ul></li><li class='menu--item'><a href='#' class='menu--link' title='Care'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Care</span></a></li><li class='menu--item'><a href='#' class='menu--link' title='Safety Day'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Safety Day</span></a></li><li class='menu--item'><a href='#' class='menu--link' title='Flynn Family'><i class='menu--icon  fa fa-fw fa-briefcase'></i><span class='menu--label'>Flynn Family</span></a></li></ul><button id='collapse_menu' class='collapse_menu'><i class='collapse_menu--icon  fa fa-fw'></i><span class='collapse_menu--label'>Recolher menu</span></button></nav>");

                            applyCSS();

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
                    RegisterModuleInit(_spPageContextInfo.siteServerRelativeUrl + 'Style Library/leftNavMenu.js', Flynn.LeftNavigation.Render.initializeHeaderFooter);
                }
                // Run now on this page (and non-MDS scenarios)
                //ExecuteOrDelayUntilScriptLoaded(Flynn.LeftNavigation.Render.initializeLeftNav, "SP.Runtime.js"); 

                Flynn.LeftNavigation.Render.initializeLeftNav();

            });
        });
    });
});


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

