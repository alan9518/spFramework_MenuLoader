// --------------------------------------
// Create Custom Flex Menu
// Production Version 1
// 4/02/2018
// --------------------------------------


// ------------------------------------------------------------------------------------------------------------------
// Get Site MetaData Module
// Production Version 1
// 4/02/2018
// ------------------------------------------------------------------------------------------------------------------
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
                $(".o365cs-nav-rightMenus").find("div:first").prepend($policyButton);
                if ($sitePolicy === "Public")
                    $("#opendg_topnavspan").html("&nbsp;&nbsp;<img src=\"https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/img/private.png\" class='ac-policyImg'>");

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
            // this.$siteURl = localStorage.getItem('hostURL');
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
        window.location.href = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" +  $sitePropsObj["_site_props_MainOwner"] + "&type=" + $sitePropsObj["_site_props_Type"] + "&cdate=" + $creationDate + "&bdate=" + $blockDate + "&comply=" + $sitePropsObj["_site_props_Compliance"] + "&curl=" + $currentUrl + "&close=" + $sitePropsObj["_site_props_closed"] + "&edate=" + $expDate + "&policy=" +  $sitePropsObj['_site_props_policy'] + "&siteid=" +  $sitePropsObj['_site_props_dbid'] + "&curuser=" + $sitePropsObj['_currentUser'];
    }

    // --------------------------------------
    // Create URL To Open on the PopUP
    // -------------------------------------
    function createURL(event)
    {
        var $creationDate = convertDate($sitePropsObj["_site_props_CreationDate"]);
        var $blockDate = convertDate($sitePropsObj["_site_props_BlockDate"]);
        var $expDate = convertDate($sitePropsObj["expDate"]);
        var $retentionURL = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + $sitePropsObj["_site_props_MainOwner"] + "&type=" + $sitePropsObj["_site_props_Type"] + "&cdate=" + $creationDate + "&bdate=" + $blockDate + "&comply=" + $sitePropsObj["_site_props_Compliance"] + "&curl=" + $currentUrl + "&close=" + $sitePropsObj["_site_props_closed"] + "&edate=" + $expDate + "&policy=" +  $sitePropsObj['_site_props_policy'] + "&siteid=" +  $sitePropsObj['_site_props_dbid'] + "&curuser=" + $sitePropsObj['_currentUser'] + "&srclink=nav&isowner=" + $sitePropsObj['_site_props_MainOwner'];
        
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

        var $extendRetentionURL = "https://flextronics365.sharepoint.com/sites/sharepoint/RetentionAssets/showmetadata.aspx?owner=" + $sitePropsObj["_site_props_MainOwner"] + "&type=" + $sitePropsObj["_site_props_Type"] + "&cdate=" + $creationDate + "&bdate=" + $blockDate + "&comply=" + $sitePropsObj["_site_props_Compliance"] + "&curl=" + $currentUrl + "&close=" + $sitePropsObj["_site_props_closed"] + "&edate=" + $expDate + "&policy=" +  $sitePropsObj['_site_props_policy'] + "&siteid=" +  $sitePropsObj['_site_props_dbid'] + "&curuser=" + $sitePropsObj['_currentUser'] + "&srclink=rtnemail&isowner=" + $sitePropsObj['_site_props_MainOwner'];

        
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

        // Show Page on the Dialog
        $($dialog).html('<iframe style="border: 0px; " src="' + $pageURL + '" width="100%" height="100%"></iframe>');

        $($dialog).dialog({
            title: "Site Governance",
            autoOpen: false,
            dialogClass: 'dialog_fixed,ui-widget-header',
            modal: true,
            width: $width,
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



// ------------------------------------------------------------------------------------------------------------------
// Create Custom Flex Menu
//Production Version 1
// 4/02/2018
// ------------------------------------------------------------------------------------------------------------------

(function()
{
    var flexMenu = 
    {
        // --------------------------------------
        // Start Module
        // --------------------------------------
        init:function()
        {
            var $menuStyles = 'https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/css/customMenuModule.css'           
            this.$spScript = 'https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/SPServices/jquery.SPServices.min.js';
            this.$userID = '';
            this.$siteURl ='';
            this.$Admins = [
                "gdljniet@americas.ad.flextronics.com",
                "jose.vazquez2@flex.com",
                "gssubhav@asia.ad.flextronics.com",
                "gssdjama@asia.ad.flextronics.com",
                "dawood.jamal@flextronics.com",
                "gdlrquin@americas.ad.flextronics.com",
                "rosa.quintero@flex.com",
                "sjcbwebb@americas.ad.flextronics.com",
                "ben.webb@flex.com",
                "gssvamar@asia.ad.flextronics.com",
                "venkatesh.amarnath@flex.com",
                "gdljacov@americas.ad.flextronics.com",
                "jannet.covarrubias@flex.com",
                "gdljaima@americas.ad.flextronics.com",
                "jaime.martinez2@flex.com",
                "gssprrav@asia.ad.flextronics.com",
                "pradeepa.ravindran@flex.com",
                "gssranbi@asia.ad.flextronics.com",
                "pedro.gamboa2@flex.com",
                "rani.anbian@flex.com",
                "luishumberto.ramirez@flex.com",
                "luis.medina@flex.com",
                "gssaekam@asia.ad.flextronics.com",
                "gssevaaa@asia.ad.flextronics.com",
                "gssgnage@asia.ad.flextronics.com",
                "gssthota@asia.ad.flextronics.com",
                "gssjpazh@asia.ad.flextronics.com",
                "gsskaven@asia.ad.flextronics.com",
                "gssmthay@asia.ad.flextronics.com",
                "gsspsuga@asia.ad.flextronics.com",
                "gsssshij@asia.ad.flextronics.com",
                "gsssrija@asia.ad.flextronics.com",
                "gdlgbena@americas.ad.flextronics.com",
                "gdllumed@americas.ad.flextronics.com",
                "gdloscac@americas.ad.flextronics.com",
                "gdlscabl@americas.ad.flextronics.com",
                "gdlswcas@americas.ad.flextronics.com",
            
                "mlgabhsh@americas.ad.flextronics.com",
                "mlgindch@americas.ad.flextronics.com", 
                "mlgasutt@americas.ad.flextronics.com",
                "mlgnaroy@americas.ad.flextronics.com",
                "mlgsochu@americas.ad.flextronics.com",
                "mlgstena@americas.ad.flextronics.com",
                "sergio.caballero@flex.com",
                "gdlswcas@americas.ad.flextronics.com",
                "swany.castrojerez@flex.com",
            
            
                "gsskveda@asia.ad.flextronics.com",
                "gdledgco@americas.ad.flextronics.com",
                "admin_gsskveda@americas.ad.flextronics.com", 
                "admin_mxedgar@asia.ad.flextronics.com",
                "admin_sacrquin@americas.ad.flextronics.com"
            
            
            ];
            
            this.loadStyles($menuStyles);
            this.isMDSEnabled();

            // console.log('admins',this.$Admins);
        
        },
        // --------------------------------------
        // Load Script and Start Process
        // --------------------------------------
        isMDSEnabled:function()
        {
            if ("undefined" != typeof g_MinimalDownload && g_MinimalDownload && (window.location.pathname.toLowerCase()).endsWith("/_layouts/15/start.aspx") && "undefined" != typeof asyncDeltaManager) {
                // Register script for MDS if possible
                RegisterModuleInit("PublicFlexBranding.js", this.RemoteManager_Inject); //MDS registration
                this.RemoteManager_Inject(); //non MDS run
            } else {
                this.RemoteManager_Inject();
            }
        },
        // --------------------------------------
        // Execute Functions
        // Load Jquery and other dependences
        // --------------------------------------
        RemoteManager_Inject:function()
        {
        var $that = this;
        if(window.jQuery )
        {    
            //  Hold Function Execution until DOM is loaded
                $.holdReady( true );
                $.getScript("https://r1.res.office365.com/o365/versionless/shellplusg2m_c305c22b.js", 
                    function()
                    {
                        $.getScript($that.$spScript,
                            function(){
                                $that.$userID = $that.getCurrentUserLoginName();
                                $that.hideFeatures();
                                $that.go();
                                // $that.createPolicyIcon();
                                getMetaDataModule.init();
                                getMetaDataModule.bindEventsPolicy();
                                $.holdReady( false );
                        })

                    });
                // add analytics
                this.executeGoogleAnalytics();
        }
        
        },
        // --------------------------------------
        // get Site Information
        // Current User
        // Use SP Services
        // --------------------------------------
        getCurrentUserLoginName:function()
        {
            var $userName = $().SPServices.SPGetCurrentUser({
                fieldNames: ["Title","EMail"],
                debug: false
            }); 

            // console.log($userName);

            return $userName.EMail;

            
        },
        // --------------------------------------
        // Hide Features
        // --------------------------------------
        hideFeatures:function()
        {
            // Look for user in admin's array
            // console.log(jQuery.inArray(this.$userID,this.$Admins))
            if(jQuery.inArray(this.$userID,this.$Admins) === -1)
            {
                this.hideSettings(false);
                this.redirectPage(false);
            }
        },
        // --------------------------------------
        // Hide Features on Settings Page
        // --------------------------------------
        hideSettings:function($isAdmin)
        {
            // console.log('Must Hide Settings');
            //For testing change for production
                if (!$isAdmin)
                {
                    $('#createnewsite').hide();
                    
                    // Settings Page
                        if ($(location).attr('href').indexOf("settings.aspx") >= 1)
                        {
                            try {
                                //Site Collection
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_DesignEditor').hide(); // Design Manager                        
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_AreaChromeSettings').hide(); // Master Page
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_AreaTemplateSettings').hide(); //  Page layouts and site templates
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_DeviceChannelSettings').hide(); // Device Channel                        
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_DesignImport').hide(); // Import Design Package
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_AreaNavigationSetting').hide(); // Navigation 
                                    $('#ctl00_PlaceHolderMain_Customization_RptControls_ImageRenditionSettings').hide(); // Image rendition


                                    $('#ctl00_PlaceHolderMain_Galleries_RptControls_MasterPageCatalog').hide(); //  Master Page
                                    $('#ctl00_PlaceHolderMain_Galleries_RptControls_CmsMasterPageCatalog').hide(); //  Master Page and page layouts
                                    $('#ctl00_PlaceHolderMain_Galleries_RptControls_Designs').hide(); // Composed Looks
                                    $('#ctl00_PlaceHolderMain_SiteTasks_RptControls_ReGhost').hide();  // Reset to site definition


                                    $('#ctl00_PlaceHolderMain_SiteTasks_RptControls_ManageSiteFeatures').hide(); // Manage Features
                                    $('#ctl00_PlaceHolderMain_SiteTasks_RptControls_SaveAsTemplate').hide(); //Save site as template
                                    $('#ctl00_PlaceHolderMain_SiteTasks_RptControls_EnableSearchConfigExport').hide(); //Enable search configuration export
                                
                                //Sub Site
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_ManageSubWebs').hide();  //Sites and workspaces
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_AreaCacheSettings').hide(); // Site output cache
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_SiteManagement').hide(); // Content and structure
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_CatalogSources').hide(); // Manage catalog connections
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_SiteManagerLogs').hide(); // Content and structure logs
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_VariationsNominateSiteLink').hide(); // Site variation settings
                                    $('#ctl00_PlaceHolderMain_SiteAdministration_RptControls_TranslationStatusListLink').hide(); // Translation Status

                            
                                //search
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_ManageResultSourcesSite').parent().parent().parent().find('.ms-linksection-title').remove(); // remove the title Search
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_ManageResultSourcesSite').hide(); // Result Sources
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_ManageResultTypes').hide(); // Result Types
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_ManageQueryRulesSite2').hide(); //Query Rules
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_MetadataPropertiesSite').hide(); //Schema
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_SiteSearchSettings').hide(); //SchemaSearch Settings
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_SrchVis').hide(); //Search and offline availability
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_SearchConfigurationImportSPWeb').hide(); //Configuration Import
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_SearchConfigurationExportSPWeb').hide(); //Configuration Export
                                    $('#ctl00_PlaceHolderMain_SearchAdministration_RptControls_NoCrawlSettingsPage').hide(); //Searchable columns
                                    
                                
                                //site collection Admin
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_DeletedItems').parent().parent().parent().find('.ms-linksection-title').remove();// remove the title Site colection Admin
                                    //$('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_DeletedItems').hide();//Recicle bin
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_ManageResultSourcesSiteColl').hide();// Search Result Resources
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_ManageResultTypes').hide();//Search result tyes
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_ManageQueryRules2').hide();//search query rules
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_MetadataPropertiesSiteColl').hide();//search schema
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_configureEnhacedSearch').hide();//search settings
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SearchConfigurationImportSPSite').hide();//search configuration Import
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SearchConfigurationExportSPSite').hide();//Search configuration Export
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_ManageSiteCollectionFeatures').hide();//Site collection features
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_Hierarchy').hide();//Site herarchy
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_AuditSettings').hide();//site collection audit settings
                                    //$('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_AuditReporting').hide();//audit log reports
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_Portal').hide();//Portal site connections
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_PolicyTemplate').hide();//content type policy templates
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SiteCollectionAppPrincipals').hide();//site collection app permissions
                                    //$('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_StorageMetrics').hide();//storage metrics
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_PolicyPolicies').hide();//Site policies
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HubUrlLinks').hide();//content type publishing
                                    //$('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SiteCollectionAnalyticsReports').hide();//popularity and search reports						
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HtmlFieldSecurity').hide();//HTML field security
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_SharePointDesignerSettings').hide();//Sharepoint Designer settings
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_HealthCheck').hide();//site collection health check
                                    $('#ctl00_PlaceHolderMain_SiteCollectionAdmin_RptControls_Upgrade').hide();	//site collection upgrade			
                                    
                                    $('#ctl00_PlaceHolderMain_UsersAndPermissions_RptControls_SiteCollectionAdministrators').hide();	//site collection administrators
                                    
                                    
                                    $('#ctl00_PlaceHolderMain_SiteTasks_RptControls_DeleteWeb').hide();  // Delete the Site 
                                
                            


                            } catch (error) {
                                alert(error.message);
                            }

                        } // Settings.aspx


                    // Site Content
                        if ($(location).attr('pathname').indexOf("viewlsts.aspx") >= 0)
                        {
                            try
                            {
                                //Create New Site link
                                $('#createnewsite').hide();
                            }
                            catch (error)
                            {
                                alert(error.message);
                            }
                        }// Site Content



                    // HIDE  FROM SITE ACTIONS
                        if ($('ie\\:menuitem[id*="_SiteActionsMenuMain_ctl00_ctl02"]').length > 0)
                            //Getting Started
                                $('ie\\:menuitem[id*="_SiteActionsMenuMain_ctl00_ctl02"]').remove();
                        
                        if ($('ie\\:menuitem[id*="SiteActionsMenuMain_ctl00_wsaDesignEditor"]').length > 0)
                            //Design Manager
                                $('ie\\:menuitem[id*="SiteActionsMenuMain_ctl00_wsaDesignEditor"]').remove();
                        
                    

                } // isAdmin
        },
        // --------------------------------------
        // Redirect Page if URL is wrotte directly
        // --------------------------------------
        redirectPage:function($isAdmin)
        {
            // console.log('Must Redirect Page');
                if (!$isAdmin)
                {                
                    // Do it only if  this is the settings web page
                    if (($(location).attr('pathname').indexOf("_layouts/15/DesignWelcomePage") >= 0) ||  // Designer
                        ($(location).attr('pathname').indexOf("_layouts/15/ChangeSiteMasterPage.aspx") >= 0) || // Master Page
                        ($(location).attr('pathname').indexOf("_layouts/15/AreaTemplateSettings.aspx") >= 0) || // Pages Layouts and site template
                        ($(location).attr('pathname').indexOf("_layouts/15/DesignPackageInstall.aspx") >= 0) || // Import Design Package
                    
                        ($(location).attr('pathname').indexOf("_layouts/15/ImageRenditionSettings.aspx") >= 0) || // Image Rendition
                        ($(location).attr('pathname').indexOf("_catalogs/masterpage") >= 0) || // Gallerires - Master Page
                        ($(location).attr('pathname').indexOf("_catalogs/design/AllItems.aspx") >= 0) || // Gallerires - Composed Looks
                        ($(location).attr('pathname').indexOf("_layouts/15/ManageFeatures.aspx") >= 0) || // Manage features
                        ($(location).attr('pathname').indexOf("_layouts/15/savetmpl.aspx") >= 0) || // Save site as template
                        ($(location).attr('pathname').indexOf("_layouts/15/Enablesearchconfigsettings.aspx") >= 0) || // Enable search configuration export
                        ($(location).attr('pathname').indexOf("_layouts/15/reghost.aspx") >= 0) || // Reset to site definition
                        ($(location).attr('pathname').indexOf("_layouts/15/mngsubwebs.aspx") >= 0) || // Sites and workspaces
                        ($(location).attr('pathname').indexOf("_layouts/15/areacachesettings.aspx") >= 0) || // Site output cache
                        ($(location).attr('pathname').indexOf("_layouts/15/sitemanager.aspx?Source={WebUrl}_layouts/15/settings.aspx") >= 0) || // Content and structure
                        ($(location).attr('pathname').indexOf("_layouts/15/ManageCatalogSources.aspx") >= 0) || // Manage catalog connections
                        ($(location).attr('pathname').indexOf("_layouts/15/SiteManager.aspx?lro=all") >= 0) || // Content and structure logs   
                        ($(location).attr('pathname').indexOf("VariationsSiteSettings.aspx") >= 0) || //Site variation settings
                        ($(location).attr('pathname').indexOf("Translation Status") >= 0) || //Translation Status 
                        ($(location).attr('pathname').indexOf("_layouts/15/newsbweb.aspx") >= 0) || //New site
                        //($(location).attr('pathname').indexOf("_layouts/15/designbuilder.aspx") >= 0) || //try out
                    
                        ($(location).attr('pathname').indexOf("manageresultsources.aspx?level=site") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("manageresulttypes.aspx?level=site") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("listqueryrules.aspx?level=site") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("listmanagedproperties.aspx?level=site") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("enhancedSearch.aspx?level=site") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("NoCrawlSettings.aspx") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("srchvis.aspx") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("importsearchconfiguration.aspx?level=site") >= 0) || //SEARCH
                        ($(location).attr('pathname').indexOf("exportsearchconfiguration.aspx?level=site") >= 0) || //SEARCH
    
                        ($(location).attr('pathname').indexOf("DeviceChannels") >= 0)   // Device Channels) 
                        //($(location).attr('pathname').indexOf("_layouts/15/AdminRecycleBin.aspx") >= 0)
                            
                        )
                    {
                        window.history.back();
                    }
                    else if ($(location).attr('pathname').indexOf('_layouts/15/deleteweb.aspx' ) >=0)
                    {
                    var $siteURl = $().SPServices.SPGetCurrentSite();
                        $siteURl += '?fromsource=compemail' ;
                        window.location.href = $siteURl;
                            
                    }
                }
        },
        // --------------------------------------
        // Start Customizing Site
        // --------------------------------------
        go:function()
        {
            var $that = this;
            try
            {
                try
                {
                    // Create Flex Menu
                    var $menuInterval  = setInterval(function(){
                        if($('#O365_MainLink_Logo').length)
                            $that.customizeMenu();
                            $('#O365_MainLink_Logo').attr("style","padding-top: 0px");
                            $('#O365_MainLink_Logo').attr("style","position: absolute");
                            // Define Menu Hover Transition
                            $that.setupMenu();
                            clearInterval($menuInterval); 
                    },400);


                    // Create Flex Footer
                    $that.createFlexFooter();
                    
                    // Show Flex Tile
                    if ($('#FlexCorpBanner').length) 
                        $('#FlexCorpBanner').append($that.getFlexTile());
                }
                catch($err)
                {
                    alert($err.message);
                }
            }
            catch($err)
            {
                alert($err.message);
            }
        },
        // --------------------------------------
        // Add Custom Navigation
        // --------------------------------------
        customizeMenu:function()
        {
            $('.o365cs-nav-o365Branding').css('position','relative');
            $('#O365_MainLink_Logo').empty();
            $('#O365_MainLink_Logo').html(

                "<a href='https://flextronics365.sharepoint.com/Pages/FlexHome.aspx'><img src='https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/img/flextrintranetlogo.png' title='Flex Home Page' height='40px' padding-top='4px' with='100%'></a>"+ 
                "<a id='mainbutton' href='#'><img src='https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/img/intranetdropdown.png' title='Global Menu'  height='40px' padding-top='4px' with='100%'></a>"+
                    "<div class='menu' style='background-color:black;position:fixed;right:0px;left:0px;z-index:999;'>"+                                                                                                            
                        "<div id='cssmenu'>"+
                            "<ul>"+
                                "<li class='has-sub'><a href='#'>Corporate</a>"+
                                "<ul id='corp' >"+
                                    "<li><a href='https://flextronics365.sharepoint.com/Pages/FlexHome.aspx'>Home</a></li>"+
                                    "<li><a href='http://news.flextronics.com/newsroom/news-overview/default.aspx'>News</a></li>"+
                                    "<li><a href='https://flextronics365.sharepoint.com/Pages/Applications.aspx'>Applications</a></li>"+
                                    "<li><a href='http://www.flextronics.com'>flex.com</a></li>"+
                                "</ul>"+
                                "</li>"+
                                "<li class='has-sub'><a href='https://flextronics365.sharepoint.com/Pages/Business%20Groups.aspx'>Business Groups</a>"+
                                    "<ul id='busg'>"+
                                        "<li><a href='https://flextronics365.sharepoint.com/ciec'>Communications &amp; Enterprise Compute</a></li>"+
                                        "<li><a href='https://flextronics365.sharepoint.com/ctg'>Consumer Technologies Group</a></li>"+
                                        "<li><a href='https://flextronics365.sharepoint.com/hrs'>High Reliability Solutions</a></li>"+
                                        "<li><a href='https://flextronics365.sharepoint.com/iei'>Industrial &amp; Emerging Industries</a></li>"+
                                        "<li><a href='https://flextronics365.sharepoint.com/innovation'>Innovation &amp; New Ventures</a></li>"+
                                    "</ul>"+
                                "</li>"+
                                "<li class='has-sub'><a href='https://flextronics365.sharepoint.com/Pages/Deparments.aspx'>Departments</a>"+
                                "<ul id='dep'>"+
                                    "<li class='deskcol'>"+ 
                                        "<a href='https://flextronics365.sharepoint.com/EQE' class='cols'>EQE – AEG, Quality, GBE</a>" +//1
                                        "<a href='https://flextronics365.sharepoint.com/sites/global_trade_compliance'  class='cols'>Global Trade</a>" +//11
                                        "<div class='clear'></div>"+ 
                                    "</li>"+
                                    "<li class='deskcol'>"+ 
                                        "<a href='https://flextronics365.sharepoint.com/components' class='cols'>Component Services</a>" +//2
                                        "<a href='https://flextronics365.sharepoint.com/hr' class='cols'>Human Resources</a>" +//12
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>" +
                                        "<a href='http://www.elementum.com' class='cols'>Elementum</a>" +//3
                                        "<a href='https://flextronics365.sharepoint.com/IT' class='cols'>Information Technology</a>" +//13
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>"+ 
                                        "<a href='https://flextronics365.sharepoint.com/EthicComp' class='cols'>Ethics & Compliance</a>" +//4
                                        "<a href='https://flextronics365.sharepoint.com/legal' class='cols'>Legal</a>" +//14
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>"+ 
                                        "<a href='https://flextronics365.sharepoint.com/finance' class='cols'>Finance</a>" +//5
                                        "<a href='https://flextronics365.sharepoint.com/marketing' class='cols'>Marketing &amp; Communications</a>" +//15
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>"+ 
                                        "<a href='https://flextronics365.sharepoint.com/GBS' class='cols'>Global Business Services</a>" +//6
                                        "<a href='https://flextronics365.sharepoint.com/multek' class='cols'>Multek</a>" +//16
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>"+   
                                        "<a href='https://flextronics365.sharepoint.com/Global%20Citizenship' class='cols'>Global Citizenship</a>" +//7
                                        "<a href='https://flextronics365.sharepoint.com/novo' class='cols'>NOVO</a>" +//17
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>" +
                                        "<a href='https://flextronics365.sharepoint.com/globalops' class='cols'>Global Operations</a>" +//8
                                        "<a href='https://flextronics365.sharepoint.com/pr' class='cols'>People and Resources</a>" +//18
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>" +
                                        "<a href='https://flextronics365.sharepoint.com/gpsc'  class='cols'>Global Procurement Supply Chain</a>" +//9
                                        "<a href='https://flextronics365.sharepoint.com/power' class='cols'>Power</a>" +//19
                                        "<div class='clear'></div>"+
                                    "</li>"+
                                    "<li class='deskcol'>" +
                                        "<a href='https://flextronics365.sharepoint.com/gss'  class='cols'>Global Services &amp Solutions</a>" +//10
                                        "<a href='https://flextronics365.sharepoint.com/strategy' class='cols'>Strategy</a>" +//20
                                        "<div class='clear'></div>"+
                                    "</li>"+
                            "</ul>"+

                            "<div class='clear'></div>"+
                            "</li>"+

                            "</ul>"+
                        "</div>"

            );

          


        },
        // --------------------------------------
        // Setup Menu Hover Transitions
        // --------------------------------------
        getFlexTile:function()
        {
        

                // Create DOM
                    var $flexTileDivParentContainer = document.createElement('div');
                        $flexTileDivParentContainer.style.width = "730px";
                    // Help
                        var $flexTileDivParentHelp = document.createElement('div');
                            $flexTileDivParentHelp.style.display = "inline-block";
                            $flexTileDivParentHelp.style.margin = "3px";

                        var $flexTileHelpLink = document.createElement('a');
                            $flexTileHelpLink.href = "https://flextronics365.sharepoint.com/sites/sharepoint/support/_layouts/15/start.aspx#/SitePages/Home.aspx";

                        var $flexTileHelpImage = document.createElement('img');
                            $flexTileHelpImage.src = "/sites/FlexSettings/SiteAssets/FlexTile/Public/NeedHelp.png";
                            $flexTileHelpImage.alt = "Learn to edit your site";
                            $flexTileHelpImage.style.border = "0";

                            $flexTileHelpLink.appendChild($flexTileHelpImage);
                            $flexTileDivParentHelp.appendChild($flexTileHelpLink);

                    // Learn

                        var $flexTileDivParentLearn = document.createElement('div');
                            $flexTileDivParentLearn.style.display = "inline-block";
                            $flexTileDivParentLearn.style.margin = "3px";

                        var $flexTileLearnLink = document.createElement('a');
                            $flexTileLearnLink.href = "http://ondemand.flextronics.com/OnDemand.html?lesson=EN_SHAREPOINT_C02_L01&flag=T";

                        var $flexTileLearnImage = document.createElement('img');
                            $flexTileLearnImage.src = "/sites/FlexSettings/SiteAssets/FlexTile/Public/Learn.png";
                            $flexTileLearnImage.alt = "Need help? find support";
                            $flexTileLearnImage.style.border = "0";
                        
                            $flexTileLearnLink.appendChild($flexTileLearnImage);
                            $flexTileDivParentLearn.appendChild($flexTileLearnLink);

                    // Permissions

                        var $flexTileDivParentPermissions = document.createElement('div');
                            $flexTileDivParentPermissions.style.display = "inline-block";
                            $flexTileDivParentPermissions.style.margin = "3px";

                        var $flexTilePermissionsLink = document.createElement('a');
                            $flexTilePermissionsLink.href = "https://flextronics365.sharepoint.com/sites/sharepoint/support/SitePages/Help.aspx?permissions=1";

                        var $flexTilePermissionsImage = document.createElement('img');
                            $flexTilePermissionsImage.src = "/sites/FlexSettings/SiteAssets/FlexTile/Public/permissions.png";
                            $flexTilePermissionsImage.alt = "Understanding Permissions";
                            $flexTilePermissionsImage.style.border = "0";

                            $flexTilePermissionsLink.appendChild($flexTilePermissionsImage);
                            $flexTileDivParentPermissions.appendChild($flexTilePermissionsLink);
                
                    // Training Videos
                        var $flexTileDivParentTraining = document.createElement('div');
                            $flexTileDivParentTraining.style.display = "inline-block";
                            $flexTileDivParentTraining.style.margin = "3px";

                        var $flexTileTrainingLink = document.createElement('a');
                            $flexTileTrainingLink.href = "http://office.microsoft.com/en-us/sharepoint-help/videos-for-sharepoint-2013-HA104071338.aspx";

                        var $flexTileTrainingImage = document.createElement('img');
                            $flexTileTrainingImage.src = "/sites/FlexSettings/SiteAssets/FlexTile/Public/videos.png";
                            $flexTileTrainingImage.alt = "Training Videos";
                            $flexTileTrainingImage.style.border = "0";


                            $flexTileTrainingLink.appendChild($flexTileTrainingImage);
                            $flexTileDivParentTraining.appendChild($flexTileTrainingLink);
                
                
                // Append DOM
                        $flexTileDivParentContainer.appendChild(flexTileDivParentHelp);    
                        $flexTileDivParentContainer.appendChild(flexTileDivParentLearn);
                        $flexTileDivParentContainer.appendChild(flexTileDivParentPermissions);
                        $flexTileDivParentContainer.appendChild(flexTileDivParentTraining);
                // Return DOM
            
                return flexTileDivParentContainer;
        },
        // --------------------------------------
        // Setup Menu Hover Transitions
        // --------------------------------------
        setupMenu:function()
        {
            var $menu = $('.menu')
            var $timeout = 0;
            var $hovering = false;
            var $menuStatus = false;
                
            $menu.hide();
            
            // Show Sub Menu
            // .searchBox_9aba68f6
                // searchBox_9aba68f6
                $('#mainbutton')
                    .on("mouseenter", function ()
                    {
                        $hovering = true;
                        // Alter Z-index Property of Search Bar
                        $('.ms-Nav').css('z-index',0);
                        var $form = $("form[role='search']")[0]
                        var $searchBox = $($form).parent();
                            $($searchBox).css('z-index',0);

                        // Open the menu
                        $('.menu')
                            .stop(true, true)
                            .slideDown(400);
                        
                
                        if ($timeout > 0)
                            clearTimeout($timeout);
                        
                    })
                    .on("mouseleave", resetHover);

            // Block SubMenu
                $('#mainbutton')
                    .on("click", function ()
                    {
                        if($menuStatus == false)
                        {
                            $hovering = true;
                            // Open the menu
                            $('.menu')
                                .stop(true, true)
                                .slideDown(400);
                            $menuStatus = true;

                            if ($timeout > 0) {
                                clearTimeout($timeout);
                            }
                        } else {
                            $hovering = false;
                            closeMenu();
                        }
                    })
            .on("mouseleave", function () {
                resetHover();
            });

            $(".menu")
                .on("mouseenter", function () {
                    // reset flag
                    $hovering = true;
                    // reset timeout
                    startTimeout();
                })
                .on("mouseleave", function () {
                    // The timeout is needed incase you go back to the main menu
                    resetHover();
                });

            function startTimeout() {
                // This method gives you 1 second to get your mouse to the sub-menu
                $timeout = setTimeout(function () {
                    closeMenu();
                }, 1000);
            };

            function closeMenu() {
                // Only close if not hovering
                if (!$hovering) {
                    $('.menu').stop(true, true).slideUp(400);
                    $menuStatus = false;
                    // Restore Z-index Property of Search Bar
                        var $form = $("form[role='search']")[0]
                        var $searchBox = $($form).parent();
                            $($searchBox).css('z-index',9);
                        $('.ms-Nav').css('z-index',8);
                }
            };

            function resetHover() {
                // Allow the menu to close if the flag isn't set by another event
                $hovering = false;
                
                // Set the timeout
                startTimeout();
            };
                
        },
        // --------------------------------------
        // Add Custom Footer
        // --------------------------------------
        createFlexFooter:function()
        {
            var $footerMessage  = 'Content on this site may not be applicable to employees of companies that have recently been acquired by Flex. Additionally, the information contained on this site is intended for Flex employees only. If you are not a Flex employee, any distribution of this content, in any form, is strictly prohibited.'
            // $('.sp-placeholder-bottom').append("<div id='flexFooter' style='text-align:center;font-family:Segoe UI light;font-size:10px;width: 1100px;position: relative;margin-left: 250px;padding-bottom:50px; '>"+ $footerMessage+"</div>");
            // Create DOM
                var $flexFooter = document.createElement('div')
                    $flexFooter.id = "flexFooter";
                    $flexFooter.style.textAlign = "center";
                    $flexFooter.style.fontFamily = "Segoe UI light";
                    $flexFooter.style.fontSize = "10px";
                    $flexFooter.style.width = "1100px";
                    $flexFooter.style.position = "relative";
                    $flexFooter.style.marginLeft = "250px";
                    $flexFooter.style.paddingBottom = "50px";
                    $flexFooter.appendChild(document.createTextNode(
                        $footerMessage
                    ));
            // Append DOM
            $('.sp-placeholder-bottom').append($flexFooter);
                                                                            
        },
        // --------------------------------------
        // Create Policy Icon
        // --------------------------------------
        createPolicyIcon:function()
        {
            var $that = this;
            this.$siteURl = $().SPServices.SPGetCurrentSite();
            if(!this.$siteURl.indexOf("https://flextronics365.sharepoint.com/sites") > -1)
                return;
            // Create Content
                var $policyIconText = document.createTextNode(this.$sitePolicy);


                var $policyIconContainer = document.createElement('div');
                    $policyIconContainer.className = 'o365cs-nav-topItem o365cs-rsp-tn-hideIfAffordanceOff';
                
            
                var $policyIconButton = document.createElement('button');
                    $policyIconButton.type = 'button';
                    $policyIconButton.className = 'o365cs-nav-item o365cs-nav-button ms-bgc-tdr-h o365button o365cs-topnavText ac-policyButton';
                    $policyIconButton.id = 'opendg_topnav';
                    $policyIconButton.setAttribute('role','menuitem');
                    $policyIconButton.setAttribute('aria-disabled','false');
                    $policyIconButton.setAttribute('aria-selected','false');
                    $policyIconButton.setAttribute('aria-label','Site Policy Type is ' + this.$sitePolicy + ' Click to change the site policy type');
                
                
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

            // Add Button to the Menu
                setTimeout(function()
                {
                    $(".o365cs-nav-rightMenus").find("div:first").prepend($policyIconContainer);
                    $("#opendg_topnavspan").html('<img src="https://stgflextronics365.sharepoint.com/sites/alan/SiteAssets/img/'+$that.$sitePolicy+'.png" class="ac-policyImg">')

                },600);

        },
        // --------------------------------------
        // Add Google Analytics
        // --------------------------------------
        executeGoogleAnalytics:function()
        {
            var dimensionValue = this.$userID;
            (function (i, s, o, g, r, a, m) {
                i['GoogleAnalyticsObject'] = r; i[r] = i[r] || function () {
                    (i[r].q = i[r].q || []).push(arguments)
                }, i[r].l = 1 * new Date(); a = s.createElement(o),
                m = s.getElementsByTagName(o)[0]; a.async = 1; a.src = g; m.parentNode.insertBefore(a, m)
            })(window, document, 'script', '//www.google-analytics.com/analytics.js', 'ga');

            if(dimensionValue != "sac_o365_moni@americas.ad.flextronics.com" && dimensionValue != "cloud_o365_moni@flextronics365.onmicrosoft.com"){
            ga('create', 'UA-56272621-1', 'auto');
            ga('set', 'dimension1', dimensionValue);
            ga('send', 'pageview');
    
        }
        },
        // --------------------------------------
        // Load CSS
        // --------------------------------------
        loadStyles:function($url)
        {
            var $flexStyles = document.createElement('link');
                $flexStyles.href = $url;
                $flexStyles.type = 'text/css';
                $flexStyles.rel  = 'stylesheet';
                $flexStyles.media = "screen,print";

            document.getElementsByTagName( "head" )[0].appendChild( $flexStyles );
        },
        // --------------------------------------
        // Load Script
        // --------------------------------------
        loadScript:function($url, $callback)
        {
            var $head =  document.getElementsByTagName('head')[0];
            var $script =  document.createElement('script');
                $script.src =  $url;
            var $done = false;
            // Attach Handlers for all Browsers
            $script.onload = $script.onreadystatechange = function()
            {
                if($done && (!this.readyState
                            ||this.readyState === 'loaded'
                            ||this.readyState === 'complete'))
                {
                    $done = true;
                    // Execute Callback Function
                    $callback();
                    // Handle Memory leak in IE
                    $script.onload = $script.onreadystatechange = null;
                    $head.removeChild($script);
                }
            };

            $head.appendChild($script);
                
        }
        
        
    }
        flexMenu.init(getMetaDataModule);
})();

