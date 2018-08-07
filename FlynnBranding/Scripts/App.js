Type.registerNamespace('Flynn');
Flynn.AddInInstall = Flynn.AddInInstall || {};

Flynn.AddInInstall.HostWebSetup = function () {
    var hostWebUrl,
        appWebUrl,
        hostWebContext,
        appWebContext,
        hostWebCustomActions,
        collUserCustomAction,
        def

        readFileFromAppWebAndProvisionToHost = function (folder,fileName) {
            $.ajax({
                url: appWebUrl + '/'+folder+'/' + fileName,
                type: 'GET',
                dataType: 'text',
                async: false,
                cache: false
            }).done(function (fileContents) {
                if (fileContents !== undefined && fileContents.length > 0) {
                    uploadFileToHostWebViaCSOM('Style Library', fileName, fileContents);
                }
                else {
                    alert('Error: Failed to read file from app web.');
                }
            });
        },

        uploadFileToHostWebViaCSOM = function (serverRelativeUrl, filename, contents) {
            var createInfo = new SP.FileCreationInformation();
            createInfo.set_content(new SP.Base64EncodedByteArray());
            for (var i = 0; i < contents.length; i++) {
                createInfo.get_content().append(contents.charCodeAt(i));
            }
            createInfo.set_overwrite(true);
            createInfo.set_url(filename);
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
        },
      deleteCustomAction = function () {
        
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
              if (oUserCustomAction.get_title() == 'jquery' || oUserCustomAction.get_title() == 'leftnavjs' || oUserCustomAction.get_title() == 'responsivecss' ||
                  oUserCustomAction.get_title() == 'normalizecss' || oUserCustomAction.get_title() == 'fontawesome') {
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

        activateLeftNav = function () {
            var oWebsite = hostWebContext.get_web();
            collUserCustomAction = oWebsite.get_userCustomActions();
            hostWebCustomActions = hostWebContext.get_web().get_userCustomActions();
            appWebContext.load(oWebsite, 'UserCustomActions', 'Title');

          
            //appWebContext.load(collUserCustomAction);
            $.when(appWebContext.executeQueryAsync(Function.createDelegate(this, this.deleteCustomAction), Function.createDelegate(this, this.onQueryFailed)))
            //deleteCustomAction();
            //hostWebCustomActions = hostWebContext.get_web().get_userCustomActions();

             .done(function (first_call, second_call, third_call) {
                 activateUserCustomAction(10023, 'jquery-1.9.1.min.js', "jquery", false);
                 activateUserCustomAction(10023, 'businessDataAccess.js', "businessdataaccess", false);
                 activateUserCustomAction(10024, 'LeftNavigationMenu.js', "leftnavjs", false);
                 // activateUserCustomAction(10025, 'vertical-responsive-menu.css',"responsivecss",true);
                  //activateUserCustomAction(10026, 'font-awesome.css',"fontawesome",true);
                 //activateUserCustomAction(10027, 'normalize.css',"normalizecss",true);

                 appWebContext.load(hostWebCustomActions);
                 appWebContext.executeQueryAsync(onActivateSuccess, onActivateError);
             });

         
        
        },
   
   

        activateUserCustomAction = function (sequence, fileName,name,css) {
            
          
            if (css) {
                var cssAction=hostWebCustomActions.add();
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
                customAction.set_scriptSrc('~SiteCollection/Style Library/' + fileName);
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

        getQueryStringParameter = function (param) {
            var params = document.URL.split("?")[1].split("&"); 

            var strParams = "";
            for (var i = 0; i < params.length; i = i + 1) {
                var singleParam = params[i].split("=");
                if (singleParam[0] == param) {
                    return singleParam[1];
                }
            }
        },

        init = function () {
            hostWebUrl = decodeURIComponent(getQueryStringParameter('SPHostUrl'));
            appWebUrl = decodeURIComponent(getQueryStringParameter('SPAppWebUrl'));

            appWebContext = new SP.ClientContext(appWebUrl);
            hostWebContext = new SP.AppContextSite(appWebContext, hostWebUrl);
        }

    return {
        execute: function () {
            init();

            // Provision these three files to the host web's Style Library and activate user custom action for each one
           
            readFileFromAppWebAndProvisionToHost('Scripts', 'LeftNavigationMenu.js');
            readFileFromAppWebAndProvisionToHost('Scripts', 'jQuery-1.9.1.min.js');
            readFileFromAppWebAndProvisionToHost('Content', 'font-awesome.css');
            readFileFromAppWebAndProvisionToHost('Content', 'normalize.css');
           readFileFromAppWebAndProvisionToHost('Content', 'vertical-responsive-menu.css');
            activateLeftNav();
        }
    }
 } ();

$(document).ready(function () {
   
    ExecuteOrDelayUntilScriptLoaded(Flynn.AddInInstall.HostWebSetup.execute, "sp.js");
});