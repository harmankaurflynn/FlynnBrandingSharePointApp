Type.registerNamespace('Flynn');
Type.registerNamespace('FlynnU');

Flynn.AddInInstall = Flynn.AddInInstall || {};

Flynn.AddInUnInstall = Flynn.AddInUnInstall || {};
FlynnU.Utility = FlynnU.Utility || {};

$(document).ready(function () {
    $("#lnkLeftNavRemove").click(function (event) {
        alert("on click");
        event.preventDefault();
        ExecuteOrDelayUntilScriptLoaded(Flynn.AddInUnInstall.HostWebRemoveSetup.execute, "sp.js");
    });

    $("#lnkActivateLeftNav").click(function (event) {
        event.preventDefault();
        ExecuteOrDelayUntilScriptLoaded(Flynn.AddInInstall.HostWebSetup.execute, "sp.js");
    });
});

Flynn.AddInUnInstall.HostWebRemoveSetup = function () {
    var hostWebUrl,
       appWebUrl,
       hostWebContext,
        hostWebRelativeURL,
       appWebContext,
       hostWebCustomActions,
       collUserCustomAction,
       def

    deactivateLeftNav = function () {
        // alert("deactivateleftnav");

        hostWebCustomActions = hostWebContext.get_web().get_userCustomActions();
        appWebContext.load(hostWebCustomActions);
        appWebContext.executeQueryAsync(onAction1Succeeded, onAction1Failed);
        ;
    },
     onAction1Succeeded = function () {
         var promise = FlynnU.Utility.Common.deleteCustomAction(appWebContext,hostWebCustomActions);
         promise.done(function () {


             deleteFilefromHostWeb('Scripts','JS', 'leftNavMenu.js');
             deleteFilefromHostWeb('Scripts', 'JS', 'jQuery-1.9.1.min.js');
             deleteFilefromHostWeb('Scripts', 'JS', 'utility.js');
             deleteFilefromHostWeb('Scripts', 'JS', 'storageManager.js');
             deleteFilefromHostWeb('Scripts', 'JS', 'businessDataAccess.js');
             deleteFilefromHostWeb('Content', 'CSS', 'font-awesome.css');
             deleteFilefromHostWeb('Content', 'CSS', 'flynnbranding.css');
             deleteFilefromHostWeb('Content', 'CSS', 'normalize.css');
             deleteFilefromHostWeb('Content', 'CSS', 'vertical-responsive-menu.css');
         })

     },
        onAction1Failed = function (sender, args) {
            alert('error');

        },
        deleteCustomActions = function () {

            //alert("deletecustomaction");

            var deleteActionPromise = [];
            // Clean up
            //foreach ( existingAction in existingActions)
            //{
            //    if (existingAction.Name.Equals(actionName, StringComparison.InvariantCultureIgnoreCase))
            //        existingAction.DeleteObject();
            //}
            //appWebContext.executeQueryAsync();


            var customActionEnumerator = hostWebCustomActions.getEnumerator();
            //var c1 = collUserCustomAction.getEnumerator();
            //collUserCustomAction.forEach((product, index) => {
            //  //  product.
            //});
            while (customActionEnumerator.moveNext()) {
                def = new $.Deferred();
                var oUserCustomAction = customActionEnumerator.get_current();
                var ocustom = oUserCustomAction;
                //alert(oUserCustomAction.get_title());
                if (oUserCustomAction.get_title() == 'jquery' || oUserCustomAction.get_title() == 'leftnavjs' || oUserCustomAction.get_title() == 'responsivecss' || oUserCustomAction.get_title() == 'storageManagerjs' ||
                    oUserCustomAction.get_title() == 'normalizecss' || oUserCustomAction.get_title() == 'fontawesome' || oUserCustomAction.get_title() == 'businessdataaccess' || oUserCustomAction.get_title() == 'utilityjs' || oUserCustomAction.get_title() == 'flynnbrandingcss') {
                    oUserCustomAction.deleteObject();
                    appWebContext.load(oUserCustomAction);
                    appWebContext.executeQueryAsync(Function.createDelegate(this, this.onQuery1Succeeded), Function.createDelegate(this, this.onQuery1Failed));
                    customActionEnumerator = hostWebCustomActions.getEnumerator();

                }
                else
                    def.resolve("ok");



            }
            return $.when.apply(undefined, deleteActionPromise).promise();
        },
     onQuery1Succeeded = function () {
         def.resolve("ok");

         // alert('Custom action removed');
     },
     onQuery1Failed = function (sender, args) {
         alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
         def.resolve("ok");
     },


        deleteFilefromHostWeb = function (folder,libFolder, fileName) {
            //init();
            // hostWebUrl = decodeURIComponent(getQueryStringParameter('SPHostUrl'));
            // appWebUrl = decodeURIComponent(getQueryStringParameter('SPAppWebUrl'));

            // appWebContext = new SP.ClientContext(appWebUrl);
            // hostWebContext = new SP.AppContextSite(appWebContext, hostWebUrl);
            $.ajax({
                url: appWebUrl + '/' + folder + '/' + fileName,
                type: 'GET',
                dataType: 'text',
                async: false,
                cache: false
            }).done(function (fileContents) {
                if (fileContents !== undefined && fileContents.length > 0) {
                    deleteFilefromLib('Style Library/FlynnBranding/'+libFolder, fileName, fileContents);
                }
                else {
                    alert('Error: Failed to read file from app web.');
                }
            });
        },

       deleteFilefromLib = function (serverRelativeUrl, filename) {



           var ofile = hostWebContext.get_web().getFileByServerRelativeUrl(hostWebRelativeURL + serverRelativeUrl + "/" + filename)
           if (ofile)
               ofile.deleteObject();

           appWebContext.executeQueryAsync(function () {


               // Success

               $('#fileDeployOutput').html('<font color=green><b>Files deleted from host web.</b></font><br/>');
           }, function (sender, args) {
               // Failure
               alert('Failed to delete files from host web. Error: ' + args.get_message());
           });




       },

    init = function () {
        hostWebUrl = decodeURIComponent(FlynnU.Utility.Common.getQueryStringParameter('SPHostUrl'));
        appWebUrl = decodeURIComponent(FlynnU.Utility.Common.getQueryStringParameter('SPAppWebUrl'));
        if (hostWebUrl.split("://")[1].split('/').length > 1)
            hostWebRelativeURL = hostWebUrl.split("://")[1].split('/')[1];
        else
            hostWebRelativeURL = "/";
        appWebContext = new SP.ClientContext(appWebUrl);
        hostWebContext = new SP.AppContextSite(appWebContext, hostWebUrl);
    }

    return {
        execute: function () {
            init();


            // Provision these three files to the host web's Style Library and activate user custom action for each one


            deactivateLeftNav();
        }
    }
}();

Flynn.AddInInstall.HostWebSetup = function () {
    var hostWebUrl,
        appWebUrl,
        hostWebContext,
        appWebContext,
        hostWebCustomActions,
        collUserCustomAction,
        hostWebRelativeURL,
        def

    readFileFromAppWebAndProvisionToHost = function (folder,libfolder, fileName) {
        $.ajax({
            url: appWebUrl + '/' + folder + '/' + fileName,
            type: 'GET',
            dataType: 'text',
            async: false,
            cache: false
        }).done(function (fileContents) {
            if (fileContents !== undefined && fileContents.length > 0) {
                uploadFileToHostWebViaCSOM('Style Library/FlynnBranding/'+libfolder, fileName, fileContents);
            }
            else {
                alert('Error: Failed to read file from app web.');
            }
        });
    },

    createFolderinhostWeb=function(serverRelativeUrl, foldername)
    {
        var deferred=new $.Deferred();
        var parentFolder = hostWebContext.get_web().getFolderByServerRelativeUrl(serverRelativeUrl);
        var ofolder = hostWebContext.get_web().getFolderByServerRelativeUrl(serverRelativeUrl + "/" + foldername);
       // if (ofolder != null)
           // parentFolder.get_folders().add(foldername);
        appWebContext.load(ofolder);

        appWebContext.executeQueryAsync(
            function () {
                deferred.resolve(ofolder);

            },
          function (sender, args) {
              ofolder = parentFolder.get_folders().add(foldername);

              appWebContext.load(ofolder);
              appWebContext.executeQueryAsync(function () {
                  deferred.resolve(ofolder);

              },
          function (sender, args) {
              deferred.reject("error");
          });

          });
            // folder exists; continue with the next part
           
                
           
      

        return deferred.promise();
    },

    uploadFileToHostWebViaCSOM = function (serverRelativeUrl, filename, contents) {
        var createInfo = new SP.FileCreationInformation();
        createInfo.set_content(new SP.Base64EncodedByteArray());
        for (var i = 0; i < contents.length; i++) {
            createInfo.get_content().append(contents.charCodeAt(i));
        }
        createInfo.set_overwrite(true);
        createInfo.set_url(filename);


        var ofile = hostWebContext.get_web().getFileByServerRelativeUrl(hostWebRelativeURL + serverRelativeUrl + "/" + filename)
        appWebContext.load(ofile);
        appWebContext.executeQueryAsync(
            Function.createDelegate(this, function () {
                /* File exists! */
                //alert('file exists');
                if (ofile.get_checkOutType() == 2) {
                    ofile.checkOut();
                    // Not checked out, do nothing
                }
                var files = hostWebContext.get_web().getFolderByServerRelativeUrl(serverRelativeUrl).get_files();
                var file = files.add(createInfo);

                appWebContext.load(file, 'CheckOutType');
                appWebContext.executeQueryAsync(function () {
                    // Success, see if file needs to be checked out
                    if (file.get_checkOutType() == 2) {
                        // Not checked out, do nothing
                    } else {
                        // Checked out
                        file.checkIn('Add-in deployment.')
                        file.publish('Add-in deployment.');
                    }

                    $('#fileDeployOutput').html('<font color=green><b>Files deployed successfully to host web.</b></font><br/>');
                }, function (sender, args) {
                    // Failure
                    alert('Failed to provision file into host web. Error: ' + args.get_message());
                });
            }),
            Function.createDelegate(this, function (sender, args) {
                /* File doesn't exist. */
                //alert(args.get_message());
                var files = hostWebContext.get_web().getFolderByServerRelativeUrl(serverRelativeUrl).get_files();
                var file = files.add(createInfo);

                appWebContext.load(file, 'CheckOutType');
                appWebContext.executeQueryAsync(function () {
                    // Success, see if file needs to be checked out
                    if (file.get_checkOutType() == 2) {
                        // Not checked out, do nothing
                    } else {
                        // Checked out
                        file.checkIn('Add-in deployment.')
                        file.publish('Add-in deployment.');
                    }

                    $('#fileDeployOutput').html('<font color=green><b>Files deployed successfully to host web.</b></font><br/>');
                }, function (sender, args) {
                    // Failure
                    alert('Failed to provision file into host web. Error: ' + args.get_message());
                });
            })
            );


    },


    activateLeftNav = function () {
        var oWebsite = hostWebContext.get_web();
        collUserCustomAction = oWebsite.get_userCustomActions();
        hostWebCustomActions = hostWebContext.get_web().get_userCustomActions();
        appWebContext.load(hostWebCustomActions);


        //appWebContext.load(collUserCustomAction);
        $.when(appWebContext.executeQueryAsync(Function.createDelegate(this, this.onActionSucceeded), Function.createDelegate(this, this.onActionFailed)));
        //deleteCustomAction();
        //hostWebCustomActions = hostWebContext.get_web().get_userCustomActions();





    },
 onActionSucceeded = function () {
     var promise = FlynnU.Utility.Common.deleteCustomAction(appWebContext,hostWebCustomActions);
     promise.then(function () {
         activateUserCustomAction(10023, 'jquery-1.9.1.min.js', "jquery", false);
         activateUserCustomAction(10024, 'utility.js', "utilityjs", false);

         activateUserCustomAction(10026, 'storageManager.js', "storageManagerjs", false);
         activateUserCustomAction(10027, 'businessDataAccess.js', "businessdataaccess", false);
         activateUserCustomAction(10028, 'leftNavMenu.js', "leftnavjs", false);


         // activateUserCustomAction(10029, 'vertical-responsive-menu.css',"responsivecss",true);
         //activateUserCustomAction(10030, 'font-awesome.css',"fontawesome",true);
         //activateUserCustomAction(10031, 'normalize.css', "normalizecss", true);
         //activateUserCustomAction(10032, 'flynnbranding.css', "flynnbrandingcss", true);

         appWebContext.load(hostWebCustomActions);
         appWebContext.executeQueryAsync(onActivateSuccess, onActivateError);
     })

 },
    onActionFailed = function (sender, args) {
        alert('error');

    },




    activateUserCustomAction = function (sequence, fileName, name, css) {


        if (css) {
            var cssAction = hostWebCustomActions.add();
            //Description: "description",
            //    customAction.set_location("ScriptLink");
            //    Name: name,
            //    customAction.set_sequence(sequence);
            //customAction.set_scriptSrc("var link=document.createElement('link'); link.rel='stylesheet'; link.type='text/css'; link.href='~SiteCollection/Style Library/' + fileName + '; document.head.appendChild(link);");


            cssAction.Location = "ScriptLink";

            cssAction.Sequence = sequence;

            cssAction.ScriptBlock = document.write("<link rel='stylesheet' href='https://dspdev05webapp.flynncompanies.com/sites/AppModelTest/Style%20Library/font-awesome.css' type='text/css' />");
            cssAction.Name = name;
            cssAction.set_title(name);

            // Apply
            cssAction.update();
            appWebContext.executeQueryAsync();



        }
        else {
            var customAction = hostWebCustomActions.add();
            customAction.set_title(name);
            customAction.set_location("ScriptLink");
            customAction.set_sequence(sequence);
            customAction.set_scriptSrc('~SiteCollection/Style Library/FlynnBranding/JS/' + fileName);
            customAction.Name = name;
            customAction.update();

        }


    },

    onActivateSuccess = function () {
        $('#activateOutput').html('<font color=green><b>Header and footer custom actions activated successfully in host web.</b></font><br/>');
    },
    onActivateError = function (sender, args) {
        alert('Failed to activate header and footer custom actions in host web. Error: ' + args.get_message());
    },



    initt = function () {
        hostWebUrl = decodeURIComponent(FlynnU.Utility.Common.getQueryStringParameter('SPHostUrl'));
        appWebUrl = decodeURIComponent(FlynnU.Utility.Common.getQueryStringParameter('SPAppWebUrl'));
        if (hostWebUrl.split("://")[1].split('/').length > 1)
            hostWebRelativeURL = hostWebUrl.split("://")[1].split('/')[1];
        else
            hostWebRelativeURL = "/";
        appWebContext = new SP.ClientContext(appWebUrl);
        hostWebContext = new SP.AppContextSite(appWebContext, hostWebUrl);
    }

    return {
        execute: function () {

            initt();
            // Provision these three files to the host web's Style Library and activate user custom action for each one
           
            var promise = createFolderinhostWeb('Style Library', 'FlynnBranding');
            promise.then(function () {
                promise = createFolderinhostWeb('Style Library/FlynnBranding', 'JS');
                promise.then(function () {
                    readFileFromAppWebAndProvisionToHost('Scripts','JS', 'leftNavMenu.js');
                    readFileFromAppWebAndProvisionToHost('Scripts', 'JS', 'jQuery-1.9.1.min.js');
                    readFileFromAppWebAndProvisionToHost('Scripts', 'JS', 'utility.js');
                    readFileFromAppWebAndProvisionToHost('Scripts', 'JS', 'storageManager.js');
                    readFileFromAppWebAndProvisionToHost('Scripts','JS', 'businessDataAccess.js');
                })
                promise = createFolderinhostWeb('Style Library/FlynnBranding', 'CSS');
                promise.then(function () {
                    readFileFromAppWebAndProvisionToHost('Content', 'CSS', 'font-awesome.css');
                    readFileFromAppWebAndProvisionToHost('Content', 'CSS', 'flynnbranding.css');
                    readFileFromAppWebAndProvisionToHost('Content', 'CSS', 'normalize.css');
                    readFileFromAppWebAndProvisionToHost('Content', 'CSS', 'vertical-responsive-menu.css');
                })
                promise = createFolderinhostWeb('Style Library/FlynnBranding', 'Fonts');
                promise.then(function () {
                    readFileFromAppWebAndProvisionToHost('Content/Font','Fonts', 'fontawesome-webfont.eot');
                   // readFileFromAppWebAndProvisionToHost('Content/Font','Fonts', 'fontawesome-webfont.ttf');
                   // readFileFromAppWebAndProvisionToHost('Content/Font','Fonts', 'fontawesome-webfont.woff');
                    //readFileFromAppWebAndProvisionToHost('Content/Font', 'Fonts', 'fontawesome-webfont.woff2');
                    readFileFromAppWebAndProvisionToHost('Content/Font', 'Fonts', 'fontawesome-webfont.svg');
                })
                activateLeftNav();
            })
        }
    }
}();

FlynnU.Utility.Common = function () {

    deleteCustomAction = function (appWebContext,hostWebCustomActions) {

        //alert("deletecustomaction");

        var deleteActionPromise = [];
        // Clean up
        //foreach ( existingAction in existingActions)
        //{
        //    if (existingAction.Name.Equals(actionName, StringComparison.InvariantCultureIgnoreCase))
        //        existingAction.DeleteObject();
        //}
        //appWebContext.executeQueryAsync();


        var customActionEnumerator = hostWebCustomActions.getEnumerator();
        //var c1 = collUserCustomAction.getEnumerator();
        //collUserCustomAction.forEach((product, index) => {
        //  //  product.
        //});
        while (customActionEnumerator.moveNext()) {
            def = new $.Deferred();
            var oUserCustomAction = customActionEnumerator.get_current();
            var ocustom = oUserCustomAction;
            //alert(oUserCustomAction.get_title());
            if (oUserCustomAction.get_title() == 'jquery' || oUserCustomAction.get_title() == 'leftnavjs' || oUserCustomAction.get_title() == 'responsivecss' || oUserCustomAction.get_title() == 'storageManagerjs' ||
                oUserCustomAction.get_title() == 'normalizecss' || oUserCustomAction.get_title() == 'fontawesome' || oUserCustomAction.get_title() == 'businessdataaccess' || oUserCustomAction.get_title() == 'utilityjs' || oUserCustomAction.get_title() == 'flynnbrandingcss') {
                oUserCustomAction.deleteObject();
                appWebContext.load(oUserCustomAction);
                appWebContext.executeQueryAsync(Function.createDelegate(this, this.onQuerySucceeded), Function.createDelegate(this, this.onQueryFailed));
                customActionEnumerator = hostWebCustomActions.getEnumerator();

            }
            else
                def.resolve("ok");



        }
        return $.when.apply(undefined, deleteActionPromise).promise();
    },
     onQuerySucceeded = function () {
         def.resolve("ok");

         // alert('Custom action removed');
     },
     onQueryFailed = function (sender, args) {
         alert('Request failed. ' + args.get_message() + '\n' + args.get_stackTrace());
         def.resolve("ok");
     },

    getQueryStringParameter = function (param) {
        var params = document.URL.split("?")[1].split("&");

        var strParams = "";
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == param) {
                return singleParam[1];
            }
        }
    }

    return {

        getQueryStringParameter: function (param) {
            return getQueryStringParameter(param);

        },
        execute: function () {
            //inith();

            // Provision these three files to the host web's Style Library and activate user custom action for each one


        },
        deleteCustomAction: function (appWebContext,hostWebCustomActions) {
            return deleteCustomAction(appWebContext,hostWebCustomActions);
        }
}
    
}();

$(document).ready(function () {

    //ExecuteOrDelayUntilScriptLoaded(Flynn.AddInInstall.HostWebSetup.execute, "sp.js");
});