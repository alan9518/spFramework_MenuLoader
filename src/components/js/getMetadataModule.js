// --------------------------------------
// Create Policy Button on Menu
// Version 1
// 2/21/2018
// --------------------------------------

var getMetaDataModule = (function ()
{

    // --------------------------------------
    // Global Variables
    // --------------------------------------
        var $sharingStatus = '';
        var $sitePolicy= "";
        var $statusBarBackground;
        var $hostUrl;
        var $subSiteUrl;
        var $currentUrl;
        var $loadPopup = "";
        var $sitePropsObj = null;

    // --------------------------------------
    // Start Module
    // --------------------------------------
    function init()
    {
        JavaScript_Embed();
    }


    // --------------------------------------
    // Bind Events
    // Use on Custom Module Menu
    // -------------------------------------
    function bindEventsPolicy()
    {
        $(document).on('click','#opendg_topnav',createURL);
        $(document).on('click','#opendgrtn',extendRetention);
        // $(document).on('click','#opendg',createURL);
    }
   
    // --------------------------------------
    // Determine URL
    // Get Values From ApplicationCustomizer.ts
    // Using LocalStorage
    // -------------------------------------
    function JavaScript_Embed()
    {
        $hostUrl = localStorage.getItem('hostURL');
        $subSiteUrl = localStorage.getItem('subSiteURL');
        if($hostUrl !== $subSiteUrl)
            $currentUrl = $subSiteUrl;
        else
            $currentUrl = $hostUrl;

        // console.log($currentUrl);
       getClassifier();
    }


    // --------------------------------------
    // Get user Name
    // -------------------------------------
    function getCurrentUserLoginName()
    {
        var $userName = $().SPServices.SPGetCurrentUser({
            fieldNames: ["Title","EMail"],
            debug: false
        }); 
        return $userName.EMail;
    }

    // --------------------------------------
    // Get Site Properties and Policy
    // -------------------------------------
    function getClassifier()
    {
        var $clientContext = SP.ClientContext.get_current();
        var $web = $clientContext.get_web();
        var $currentUser = $web.get_currentUser();
        var $props = $web.get_allProperties();
            $clientContext.load($web);
            $clientContext.load($currentUser);
            $clientContext.load($web, 'EffectiveBasePermissions');
            $clientContext.load($props);

            $clientContext.executeQueryAsync(function (sender, args)
            {
                try
                {
                    var $isOwner = "No";
                    if ($web.get_effectiveBasePermissions().has(SP.PermissionKind.manageWeb))
                        $isOwner = "Yes"; //User Has Edit Permissions
                        
                    
                    // Check Policy
                    if($sitePolicy.length !== 0)
                    {
                        $sitePropsObj = 
                        {
                            _site_props_MainOwner :        $props.get_item("_site_props_MainOwner"),
                            _site_props_Type :             $props.get_item("_site_props_Type"),
                            _site_props_CreationDate :     $props.get_item("_site_props_CreationDate"),
                            _site_props_BlockDate :        $props.get_item("_site_props_BlockDate"),
                            _site_props_Compliance :       $props.get_item("_site_props_Compliance"),
                            _site_props_SharingCapability: $props.get_item('_site_props_SharingCapability')
                        };


                        // var $sharingStatus1 = $props.get_item("_site_props_MainOwner");
                        // var $sharingStatus2 = $props.get_item("_site_props_Type");
                        // var $sharingStatus3 = $props.get_item("_site_props_CreationDate");
                        // var $sharingStatus4 = $props.get_item("_site_props_BlockDate");
                        // var $sharingStatus5 = $props.get_item("_site_props_Compliance");

                        // getSiteSharingStatus();
                        console.log('Obj',$sitePropsObj)
                    }
                    else
                    {
                        $sitePolicy= "No Site Policy";
                        $sitePropsObj = 
                        {
                            _site_props_MainOwner :        $props.get_item("_site_props_MainOwner"),
                            _site_props_Type :             $props.get_item("_site_props_Type"),
                            _site_props_CreationDate :     $props.get_item("_site_props_CreationDate"),
                            _site_props_BlockDate :        $props.get_item("_site_props_BlockDate"),
                            _site_props_Compliance :       $props.get_item("_site_props_Compliance"),
                            _site_props_closed :           $props.get_item("_site_props_closed"),
                            _site_props_expirydate :       $props.get_item("_site_props_expirydate"),
                            _site_props_policy :           $props.get_item("_site_props_policy"),
                            _site_props_dbid :             $props.get_item("_site_props_dbid"),
                            _site_props_SharingCapability: $props.get_item('_site_props_SharingCapability'),
                            _currentUser :                  getCurrentUserLoginName()
                        };

                        // var $sharingStatus1 = $props.get_item("_site_props_MainOwner");
                        // var $sharingStatus2 = $props.get_item("_site_props_Type");
                        // var $sharingStatus3 = $props.get_item("_site_props_CreationDate");
                        // var $sharingStatus4 = $props.get_item("_site_props_BlockDate");
                        // var $sharingStatus5 = $props.get_item("_site_props_Compliance");
                        // var $sharingStatus6 = $props.get_item("_site_props_closed");
                        // var $sharingStatus7 = $props.get_item("_site_props_expirydate");
                        // var $sharingStatus8 = $props.get_item("_site_props_policy");
                        // var $sharingStatus9 = $props.get_item("_site_props_dbid");
                        // var $currentUser = getCurrentUserLoginName();

                        // getSiteSharingStatus();

                        
                        // console.log('Obj',$sitePropsObj)
                        setUI($sitePropsObj);
                    }
                }
                catch(e)
                {
                    $sitePolicy= "No Site Policy";
                    $sitePropsObj = 
                        {
                            _site_props_MainOwner :        $props.get_item("_site_props_MainOwner"),
                            _site_props_Type :             $props.get_item("_site_props_Type"),
                            _site_props_CreationDate :     $props.get_item("_site_props_CreationDate"),
                            _site_props_BlockDate :        $props.get_item("_site_props_BlockDate"),
                            _site_props_Compliance :       $props.get_item("_site_props_Compliance"),
                            _site_props_closed :           $props.get_item("_site_props_closed"),
                            _site_props_expirydate :       $props.get_item("_site_props_expirydate"),
                            _site_props_policy :           $props.get_item("_site_props_policy"),
                            _site_props_dbid :             $props.get_item("_site_props_dbid"),
                            _site_props_SharingCapability: $props.get_item('_site_props_SharingCapability'),
                            _currentUser :                  getCurrentUserLoginName()
                        };
                    // var $sharingStatus1 = $props.get_item("_site_props_MainOwner");
                    // var $sharingStatus2 = $props.get_item("_site_props_Type");
                    // var $sharingStatus3 = $props.get_item("_site_props_CreationDate");
                    // var $sharingStatus4 = $props.get_item("_site_props_BlockDate");
                    // var $sharingStatus5 = $props.get_item("_site_props_Compliance");
                    // var $sharingStatus6 = $props.get_item("_site_props_closed");
                    // var $sharingStatus7 = $props.get_item("_site_props_expirydate");
                    // var $sharingStatus8 = $props.get_item("_site_props_policy");
                    // var $sharingStatus9 = $props.get_item("_site_props_dbid");
                    // var $currentUser = getCurrentUserLoginName();

                    // getSiteSharingStatus();

                    setUI($sitePropsObj);
                        
                    // console.log('Obj',$sitePropsObj)
                }
            })
    }
    

    // --------------------------------------
    // Create UI
    // -------------------------------------
    function setUI($sitePropsObj)
    {
            $sitePolicy = $sitePropsObj["_site_props_policy"]
        var $policyButton  = createPolicyIcon();   
        var $currentDate = new Date();
            $currentDate.setHours(0,0,0,0);
        var $expDate = new Date($sitePropsObj["_site_props_expirydate"]);
            $expDate.setHours(0,0,0,0);
        
        var $diffDays = parseInt(($expDate - $currentDate) / (1000 * 60 * 60 * 24));

        // Add Image
            setTimeout(function()
            {
                // $(".o365cs-nav-rightMenus").find("div:first").prepend($policyButton);
                // if ($sitePolicy === "Public")
                //     $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/img/public.png\" class='ac-policyImg'>");
                // else 
                //     $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/img/private.png\" class='ac-policyImg'>");

                $(".o365cs-nav-rightMenus").find("div:first").prepend($policyButton);
                if ($sitePolicy === "Public")
                    $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/img/public.png\" class='ac-policyImg'>");
                else 
                    $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/img/private.png\" class='ac-policyImg'>");

            },300);

        // set Message and status Bar Background

            $sitePropsObj['expDate'] = $expDate;
            $sitePropsObj['currentDate'] = $currentDate;
            $sitePropsObj['diffDays'] = $diffDays;
            setMessage($sitePropsObj);
    }


    // --------------------------------------
    // Create the Button
    // -------------------------------------
    function createPolicyIcon()
    {
            var $that = this;
            // this.$siteURl = $().SPServices.SPGetCurrentSite();
            // if(!this.$siteURl.indexOf("https://flextronics365.sharepoint.com/sites") > -1)
            //     return;
            // else
        // Create Content
                var $policyIconText = document.createTextNode($sitePolicy);


                var $policyIconContainer = document.createElement('div');
                    $policyIconContainer.className = 'o365cs-nav-topItem o365cs-rsp-tn-hideIfAffordanceOff';
                
            
                var $policyIconButton = document.createElement('button');
                    $policyIconButton.type = 'button';
                    $policyIconButton.className = 'o365cs-nav-item o365cs-nav-button ms-bgc-tdr-h o365button o365cs-topnavText ac-policyButton';
                    $policyIconButton.id = 'opendg_topnav';
                    $policyIconButton.setAttribute('role','menuitem');
                    $policyIconButton.setAttribute('aria-disabled','false');
                    $policyIconButton.setAttribute('aria-selected','false');
                    $policyIconButton.setAttribute('aria-label','Site Policy Type is ' +$sitePolicy + ' Click to change the site policy type');
                var $linkAtt = document.createAttribute('data-link');
                    $policyIconButton.setAttributeNode($linkAtt);
                

                
                var $policyIconSpan = document.createElement('span');
                    $policyIconSpan.id = 'opendg_topnavspan';
                    $policyIconSpan.className = 'o365cs-topnavText owaimg ms-Icon--policy ms-icon-font-size-20';
                    $policyIconSpan.setAttribute('aria-hidden','true');
                
                    
                
               
                var $policyIconUnSeenContainer = document.createElement('div');
                    $policyIconUnSeenContainer.className = 'o365cs-flexPane-unseenitems';
                
                var $policyIconUnSeenSpan =  document.createElement('span');
                    $policyIconUnSeenSpan.className = 'o365cs-flexPane-unseenCount ms-bgc-tdr ms-fcl-w';
                    $policyIconUnSeenSpan.style.display = 'none';
            
                var $policyIconUnSeenSpanCount =  document.createElement('span');
                    $policyIconUnSeenSpanCount.className = 'o365cs-flexPane-unseenCount owaimg ms-Icon--starburst ms-icon-font-size-12 ms-bgc-tdr ms-fcl-w';
                    $policyIconUnSeenSpanCount.style.display = 'none';



            // Append Items of the Button
                $policyIconButton.appendChild($policyIconSpan);
                $policyIconButton.appendChild($policyIconText);

                $policyIconUnSeenContainer.appendChild($policyIconUnSeenSpan);
                $policyIconUnSeenContainer.appendChild($policyIconUnSeenSpanCount);

                $policyIconContainer.appendChild($policyIconButton);
                $policyIconContainer.appendChild($policyIconUnSeenContainer);
            
            // Return Button
            return ($policyIconContainer);
    }
 

   
    // --------------------------------------
    // set Message and status Bar Background
    // Using site Policy
    // --------------------------------------
    function setMessage($sitePropsObj)
    {
        var $policy = $sitePropsObj['_site_props_policy'];
        var $compliance = $sitePropsObj['_site_props_Compliance'];
            $sharingStatus =  $sitePropsObj['_site_props_SharingCapability'];
        var $closed = $sitePropsObj['_site_props_closed'];
        var $diffDays =  $sitePropsObj['diffDays'];
        var $message = '';
        // High Business Impact
            if ($policy === 'High Business Impact')
            {
                if ($sharingStatus === "true" || $sharingStatus === "True") 
                    $message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>High Business Impact</b> and <b>Partner Sharing is enabled</b></a></font>";
                else 
                    $message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>High Business Impact</b></a></font>";
                $statusBarBackground = "#f0f0f0";
            }
        
        // Medium Business Impact
            else if ($policy == 'Medium Business Impact')
            {
                if ($sharingStatus === "true" || $sharingStatus === "True")
                    $message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Medium Business Impact</b> and <b>Partner Sharing is enabled</b></a></font>";
                else
                    $message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Medium Business Impact</b></a></font>";
                $statusBarBackground = "#f0f0f0";
            }

        // Low Business Impact
            else if ($policy == 'Low Business Impact')
            {
                if ($sharingStatus == "true" || $sharingStatus == "True")
                    $message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Low Business Impact</b> and <b>Partner Sharing is enabled</b></a></font>";
                else
                    $message = "<font color='#000000'><a href='https://yoursitepolicy' target=_blank>Information on this site has been classified as <b>Low Business Impact</b></a></font>";
                $statusBarBackground = "#f0f0f0";
            }    

        // Site Expired diffDays < 0
            else if ($diffDays < 0)
            {
                if ($sharingStatus === "true" || $sharingStatus === "True")
                    $message = "<font color='#000000'>Information on this site has not yet been classified.  Click <a href='#' id='opendg'>here</a> to set the Policy. <b>Partner sharing is enabled</b></font>";
                else
                    $message = "<font color='#000000'>This site looks already expired.  Click <a href='#' id='opendgrtn'>here</a> to extend the site expiry date and to keep this site open.</font>";
                $statusBarBackground = "red";
            }
        // Site About to Expire on the last 30 Days
            else if ($diffDays >= 0 && $diffDays <= 30)
            {
                if ($sharingStatus === "true" || $sharingStatus === "True") 
                    $message = "<font color='#000000'>Information on this site has not yet been classified.  Click <a href='#' id='opendgrtn'>here</a> to set the Policy. <b>Partner sharing is enabled</b></font>";
                else 
                    $message = "<font color='#000000'>This site will be expired with in " + $diffDays + " days .  Click <a href='#' id='opendgrtn'>here</a> to extend the site expiry date.</font>";
                $statusBarBackground = "red";
            }   

        // No Compliance Found
            else if ($compliance === "false") 
            {
                if ($sharingStatus === "true" || $sharingStatus === "True") 
                    $message = "<font color='#000000'>Information on this site has not yet been classified.  Click <a href='#' id='opendg'>here</a> to set the Policy. <b>Partner sharing is enabled</b></font>";
                else 
                    $message = "<font color='#000000'>Site is not under compliance based on the SharePoint site governance.  Click <a href='#' id='opendg'>here</a> to validate the site compliance.</font>";
                $statusBarBackground = "yellow";
            }
        
        // Site is Closed
            if ($closed == "true") 
                // window.location.href = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + MainOwner + "&type=" + Type + "&cdate=" + CreationDate + "&bdate=" + BlockDate + "&comply=" + Compliance + "&curl=" + currentUrl + "&close=" + closed + "&edate=" + expdate + "&policy=" + policy + "&siteid=" + siteid+"&curuser="+curuser;
                redirectSite($sitePropsObj);
        
        // Add code to set a policy (reminder red) This sub-site does not have a policy set. Click here to set
            if ($message.length > 0)
                SetStatusBar($message, $statusBarBackground);
            

    }

    // --------------------------------------
    // get Site Sharing Status
    // -------------------------------------
    function getSiteSharingStatus()
    {   
        if ($hostUrl == $subSiteUrl) {
            var $clientContext = SP.ClientContext.get_current();
            var $web = $clientContext.get_web();
            var $props = $web.get_allProperties();
                $clientContext.load($web);
                $clientContext.load($props);
    
            $clientContext.executeQueryAsync(function (sender, args)
            {
                $sharingStatus = props.get_item('_site_props_SharingCapability');
                setUI();
    
            });
    
            return $sharingStatus;
        }
        else {
            var $clientContext = SP.ClientContext.get_current();
            var $site = $clientContext.get_site();
            var $web = site.get_rootWeb();
            var props = $web.get_allProperties();
                $clientContext.load($web);
                $clientContext.load($props);
        
            $clientContext.executeQueryAsync(function (sender, args)
            {
                $sharingStatus = props.get_item('_site_props_SharingCapability');
                //var $sharingStatus1 = props.get_item("_site_props_MainOwner");
                //var $sharingStatus2 = props.get_item("_site_props_Type");
                //var $sharingStatus3 = props.get_item("_site_props_CreationDate");
                //var $sharingStatus4 = props.get_item("_site_props_BlockDate");
                //var $sharingStatus5 = props.get_item("_site_props_Compliance");
                //var $sharingStatus6 = props.get_item("_site_props_closed");
                //var $sharingStatus7 = props.get_item("_site_props_expirydate");
                //var $sharingStatus8 = props.get_item("_site_props_policy");	
                
                
                //setUI($sharingStatus1,$sharingStatus2,$sharingStatus3,$sharingStatus4,$sharingStatus5);
    
            });
    
            return $sharingStatus;
        }
        
    }

   

    // --------------------------------------
    // Redirect Site
    // -------------------------------------
    function redirectSite($sitePropsObj)
    { 
        window.location.href = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" +  $sitePropsObj["_site_props_MainOwner"] + "&type=" + $sitePropsObj["_site_props_Type"] + "&cdate=" + $creationDate + "&bdate=" + $blockDate + "&comply=" + $sitePropsObj["_site_props_Compliance"] + "&curl=" + $currentUrl + "&close=" + $sitePropsObj["_site_props_Compliance"] + "&edate=" + $expDate + "&policy=" +  $sitePropsObj['_site_props_policy'] + "&siteid=" +  $sitePropsObj['_site_props_dbid'] + "&curuser=" + $sitePropsObj['_currentUser'];
    }

    // --------------------------------------
    // Create URL To Open on the PopUP
    // -------------------------------------
    function createURL(event)
    {
        var $creationDate = convertDate($sitePropsObj["_site_props_CreationDate"]);
        var $blockDate = convertDate($sitePropsObj["_site_props_BlockDate"]);
        var $expDate = convertDate($sitePropsObj["expDate"]);
        var $retentionURL = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + $sitePropsObj["_site_props_MainOwner"] + "&type=" + $sitePropsObj["_site_props_Type"] + "&cdate=" + $creationDate + "&bdate=" + $blockDate + "&comply=" + $sitePropsObj["_site_props_Compliance"] + "&curl=" + $currentUrl + "&close=" + $sitePropsObj["_site_props_Compliance"] + "&edate=" + $expDate + "&policy=" +  $sitePropsObj['_site_props_policy'] + "&siteid=" +  $sitePropsObj['_site_props_dbid'] + "&curuser=" + $sitePropsObj['_currentUser'] + "&srclink=nav&isowner=" + $sitePropsObj['_site_props_MainOwner'];
        
        // console.log('Retention URL',$retentionURL);
        openDialog($retentionURL);
    }

    // --------------------------------------
    // Extend Site Retention URL
    // -------------------------------------
    function extendRetention()
    {
        var $creationDate = convertDate($sitePropsObj["_site_props_CreationDate"]);
        var $blockDate = convertDate($sitePropsObj["_site_props_BlockDate"]);
        var $expDate = convertDate($sitePropsObj["expDate"]);

        var $extendRetentionURL = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + $sitePropsObj["_site_props_MainOwner"] + "&type=" + $sitePropsObj["_site_props_Type"] + "&cdate=" + $creationDate + "&bdate=" + $blockDate + "&comply=" + $sitePropsObj["_site_props_Compliance"] + "&curl=" + $currentUrl + "&close=" + $sitePropsObj["_site_props_Compliance"] + "&edate=" + $expDate + "&policy=" +  $sitePropsObj['_site_props_policy'] + "&siteid=" +  $sitePropsObj['_site_props_dbid'] + "&curuser=" + $sitePropsObj['_currentUser'] + "&srclink=rtnemail&isowner=" + $sitePropsObj['_site_props_MainOwner'];

        
    }

    // --------------------------------------
    // Open JQuery UI Dialog
    // --------------------------------------
    function openDialog($pageURL)
    {
        var $width = window.innerWidth;
        var $height = window.innerHeight;
            $width = $width/1.6;
            $height = $height/1.2;

        // Update Current Dialog Content
        if( $('#modalRetention').length > 0 )
        {
            $($dialog).empty();
            $($dialog).html('<iframe style="border: 0px; " src="' + $pageURL + '" width="100%" height="100%"></iframe>');
            return;
        }
        // Create New Dialog
        var $dialog = document.createElement('div');
            $dialog.id ='modalRetention';
            $dialog.className ='modalRetention';

        $($dialog).html('<iframe style="border: 0px; " src="' + $pageURL + '" width="100%" height="100%"></iframe>');

        $($dialog).dialog({
            title: "Site Governance",
            autoOpen: false,
            dialogClass: 'dialog_fixed,ui-widget-header',
            modal: true,
            width:$width,
            height: $height,
            minWidth: $width,
            minHeight: $height,
            closeText: 'X',
            draggable:true,
            close: function () { $(this).remove(); },
          });

       $($dialog).dialog('open');
    }   


    // --------------------------------------
    // Convert Full Date to 
    // DD/MM/YYYY
    // -------------------------------------
    function convertDate($date)
    {
        var $nDate = new Date($date);
        return ($nDate.getFullYear()+ "/" + ($nDate.getMonth() + 1) + "/" + $nDate.getDate());
    }

    // --------------------------------------
    // set Status Bar
    // -------------------------------------
    function SetStatusBar($message, $bgColor) 
    {
        var $strUpdatedStatusID = SP.UI.Status.addStatus("Attention: ", $message, true);
        SP.UI.Status.setStatusPriColor($strUpdatedStatusID, $bgColor);    
    }

    // --------------------------------------
    // Get URL Parameters
    // -------------------------------------
    function getParameterByName(name)
    {
        try
        {
            var $params = document.URL.split("?")[1].split("&");
            var $strParams = "";
            for (var i = 0; i < $params.length; i = i + 1)
            {
                var $singleParam = $params[i].split("=");
                if ($singleParam[0] == name)
                    return $singleParam[1];
            }
        }
        catch(err)
        {
            return "null";
        }
    }

    // Reveal public pointers to
    // private functions and properties
    // To Other Modules
    return {
        init:init,
        bindEventsPolicy: bindEventsPolicy
    };

    


}) 
();



