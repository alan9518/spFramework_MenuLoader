// --------------------------------------
// Create Custom Flex Menu
// Version 1.2
// 2/21/2018
// --------------------------------------

(function()
{
    var flexMenu = 
    {
        // --------------------------------------
        // Start Module
        // --------------------------------------
        init:function()
        {
            var $menuStyles = 'https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/css/customMenuModule.css'           
            this.$spScript = 'https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/SPServices/jquery.SPServices.min.js'
            this.$metaDataScript = 'https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/js/getMetadataModule.js'
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
        // Cache Dom
        // --------------------------------------
        // cacheDOM:function()
        // {
        //     this.$mainMenu = $('#SuiteNavPlaceHolder');
        // },
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
            console.log('Must Hide Settings');
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
            console.log('Must Redirect Page');
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
                        $('#FlexCorpBanner').html($that.getFlexTile());
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
            $('#O365_MainLink_Logo').empty();
            $('#O365_MainLink_Logo').html(

                "<a href='https://flextronics365.sharepoint.com/Pages/FlexHome.aspx'><img src='https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/img/flextrintranetlogo.png' title='Flex Home Page' height='40px' padding-top='4px' with='100%'></a>"+ 
                "<a id='mainbutton' href='#'><img src='https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/img/intranetdropdown.png' title='Global Menu'  height='40px' padding-top='4px' with='100%'></a>"+
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
            var strRestul = "";

            strRestul =`<div style='width:730px;'>
                            <div style='display:inline-block; margin: 3px;' > 
                            <a href='https://flextronics365.sharepoint.com/sites/sharepoint/support/_layouts/15/start.aspx#/SitePages/Home.aspx' >
                            <img style='border:0' src='/sites/FlexSettings/SiteAssets/FlexTile/Public/NeedHelp.png' alt='Learn to edit your site' /> 
                            </a>
                            </div>
                            <div style='display:inline-block;margin: 3px;'> 
                            <a href='http://ondemand.flextronics.com/OnDemand.html?lesson=EN_SHAREPOINT_C02_L01&flag=T' >
                            <img style='border:0' src='/sites/FlexSettings/SiteAssets/FlexTile/Public/Learn.png' alt='Need help? find support' />
                            </a>
                            </div>
                            <div style='display:inline-block;margin: 3px;'> 
                            <a href='https://flextronics365.sharepoint.com/sites/sharepoint/support/SitePages/Help.aspx?permissions=1' >
                            <img style='border:0' src='/sites/FlexSettings/SiteAssets/FlexTile/Public/permissions.png' alt='Understanding Permissions' /> 
                            </a>
                            </div>
                            <div style='display:inline-block;margin: 3px;'> 
                            <a href='http://office.microsoft.com/en-us/sharepoint-help/videos-for-sharepoint-2013-HA104071338.aspx' >
                            <img style='border:0' src='/sites/FlexSettings/SiteAssets/FlexTile/Public/videos.png' alt='Training Videos' /> 
                            </a>
                            </div>
                        </div>`;
                
            return strRestul;
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
            $('.sp-placeholder-bottom').append("<div id='flexFooter' style='text-align:center;font-family:Segoe UI light;font-size:10px;width: 1100px;position: relative;margin-left: 250px;padding-bottom:50px; '>"+ $footerMessage+"</div>");
                                                                              
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
                    $("#opendg_topnavspan").html('<img src="https://flextronics365.sharepoint.com/sites/FlexSettings/Style%20Library/custom%20menu/img/'+$that.$sitePolicy+'.png" class="ac-policyImg">')

                },600);

        },
        openPolicyWindow:function()
        {
            console.log('click button')
        },
        // --------------------------------------
        // Get Current User LoginName
        // --------------------------------------
        getXMLDataPromise:function($serviceURL)
        {
            // Create Promise
            return new Promise(function(resolve,reject){

                // DO XHR Stuff
                var $req = new XMLHttpRequest();
                    $req.open('GET',$serviceURL);

                    $req.onload = function()
                    {
                        // Check Answer Status
                        if($req.status === 200)
                        {
                            // Resolve the Promise with the response Text
                            resolve($req.response);
                        }
                        else
                        {
                            // Otherwise reject with the status text
                            // which will hopefully be a meaningful error
                            reject(Error($req.statusText ));
                        }
                    };

                    // Handle Network Error
                    $req.onerror = function()
                    {
                        reject(Error("Network error"));
                    }
                
                // Make the Request
                $req.send();
            });
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
        // flexMenu.init();
})();

