// -------------------------------------------------------------------------------
//  
//   GCP Project Number: seo-data 382643835883
//
// -------------------------------------------------------------------------------
var app = {
  APP_TITLE: "Sequencer Playlists",
  MENU_TITLE_CREATE: "Create Playlists",
  MENU_TITLE_PARK: "Park Playlists",
  MENU_TITLE_RESTORE: "Restore Playlists",
  MENU_TITLE_VIEW: "View Playlists",
  PLAYLIST_RECORD_NAME: "playlists",
  PLAYLIST_RECORD_NAME_PARK: 'playlists_parked',
  PLAYLIST_TABLE_NAME: "",
  DATASTORE_TABLE_ROWCNT: 2,
  DATASTORE_TABLE_COLCNT: 2,
  DEFAULT_DATA_MODEL: { 'pl' :[], 'slideMeta': [] },
  DEFAULT_THUMBNAIL_URL: 'https://i.pcmag.com/imagery/reviews/03ErPVuqnBDCwlLsh8EzpBM-5.fit_scale.size_1028x578.v_1569477508.jpg'
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
var viewModel = {
  count:0,
  slideMeta: [ ]
}

var thumbNailCache = {
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
viewModel.updateThumbnailCache = function(slideId) { 
  slideThumbTask = new Promise(function(completeTask, failTask){
    var parameters = {
      method: "GET",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      contentType: "application/json",
      muteHttpExceptions: true
    };
    var httpUrl = "https://slides.googleapis.com/v1/presentations/" + SlidesApp.getActivePresentation().getId() + "/pages/" + slideId + "/thumbnail"
    var response = UrlFetchApp.fetch(httpUrl, parameters);
    completeTask({ id: slideId, urlData: JSON.parse(response.getContentText()) })
  }).then((a) => {
    setDataString(a.id, a.urlData.contentUrl)
  })
}

function test() {
  viewModel.updateThumbnailCache('g72b54a5b54_0_18')
}

function qwerty(i) {
  return { url:getDataString(i), id:i }
}


function getCachedUrl(id) {
  var o = qwerty(id)
  if(o != null) {
    return o.url
  }
  return ''
}

//
// Create an object array for each each image
//
function bobbins( pl ) {
  for (var i = 0; i < pl.length; i++) {
    var plName = pl[i].name 
    for (var j = 0; j < pl[i].ids.length; j++) {
      var thisSlideId = pl[i].ids[j]
    }
  }
}


// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
viewModel.getThumbnailPathForId = function(slideId) { 
  var url = thumbNailCache[slideId]
  if(url == null) {
    url = app.DEFAULT_THUMBNAIL_URL  
  }  
  return url
}

viewModel.cleanMeta = function() { 
}

viewModel.buildSideBar = function() { 
}

// -------------------------------------------------------------------------------
//
// OnInstall: build the meui items and other relevant UI assets
//
// -------------------------------------------------------------------------------
function onInstall(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem(app.MENU_TITLE_CREATE, 'onCreate')
      .addItem(app.MENU_TITLE_VIEW, 'showSideBar')
      .addItem(app.MENU_TITLE_PARK, 'parkDataStore')
      .addItem(app.MENU_TITLE_RESTORE, 'restoreDataStore')
      .addToUi()
      initDataStore()
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem(app.MENU_TITLE_CREATE, 'onCreate')
      .addItem(app.MENU_TITLE_VIEW, 'showSideBar')
      .addItem(app.MENU_TITLE_PARK, 'parkDataStore')
      .addItem(app.MENU_TITLE_RESTORE, 'restoreDataStore')
      .addToUi()
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function refreshUserProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('loopCounter', 0);
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function showSideBar() { 

  try {
    var playListSideBarHTML = createSidebarHTML()
    var htmlOutput = HtmlService.createHtmlOutput(playListSideBarHTML).setTitle(app.APP_TITLE);
    SlidesApp.getUi().showSidebar(htmlOutput);
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }  
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function runPlayList(plId) {
  try {
    SlidesApp.getUi().alert('running ...' + plId)
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function deletePlayList(name) {
  try {
     var playLists = getData(app.PLAYLIST_RECORD_NAME)
     var updatedPlaylist = playLists.pl.filter(function(value, index, arr){ return value.name != name;});
     playLists.pl = updatedPlaylist
     setData(app.PLAYLIST_RECORD_NAME, playLists)
     showSideBar();
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
}

// -------------------------------------------------------------------------------
//
// Send the slides in a playlist email attachments
//
// -------------------------------------------------------------------------------
function mailPlaylistSlides(name) {
  //var file = DriveApp.getFileById('1234567890abcdefghijklmnopqrstuvwxyz');
  //var blob = Utilities.newBlob('Insert any HTML content here', 'text/html', 'my_document.html');
  //MailApp.sendEmail('mike@example.com', 'Attachment example', 'Two files are attached.', {
  //    name: 'Automatic Emailer Script',
  //    attachments: [file.getAs(MimeType.PDF), blob]
  //});
}

// -------------------------------------------------------------------------------
//
// Package slides in a playlist as PDF attachment
//
// -------------------------------------------------------------------------------
function mailPDFPlaylistSlides(name) {
  //var file = DriveApp.getFileById('1234567890abcdefghijklmnopqrstuvwxyz');
  //var blob = Utilities.newBlob('Insert any HTML content here', 'text/html', 'my_document.html');
  //MailApp.sendEmail('mike@example.com', 'Attachment example', 'Two files are attached.', {
  //    name: 'Automatic Emailer Script',
  //    attachments: [file.getAs(MimeType.PDF), blob]
  //});
}

/*
  // Log URL of the main thumbnail of the deck
  //Logger.log(Drive.Files.get(presentationId).thumbnailLink);

  // For storing the screenshot image URLs
  //var screenshots = [];

  //var slides = presentation.getSlides().forEach(function(slide, index) {
  //  var url = baseUrl
  //    .replace("{presentationId}", presentationId)
  //    .replace("{pageObjectId}", slide.getObjectId());
  //  var response = JSON.parse(UrlFetchApp.fetch(url, parameters));

    // Upload Googel Slide image to Google Drive
  //  var blob = UrlFetchApp.fetch(response.contentUrl).getBlob();
  //  DriveApp.createFile(blob).setName("Image " + (index + 1) + ".png");

  //  screenshots.push(response.contentUrl);
 // });  
*/


// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function playFromPlaylist(idArray) {
  try {
     for(var slideId in idArray) {
       selectSlide(slideId)
       break; // Utilities.sleep(1000)
     }
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function play(slideArray) {
  if(slideArray == null || slideArray.length < 1) {
    return
  }
  else {

    //Utilities.sleep(500)
    //play(slideArray)
  }
  return 
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function doNothing(plId) {
  try {
    SlidesApp.getUi().alert('running ...' + plId)
  }
  catch (e) {
  SlidesApp.getUi().alert(e.message)
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function selectSlide(plId) {
  try {
    SlidesApp.getActivePresentation().getSlideById(plId).selectAsCurrentPage()
  }
  catch (e) {
  SlidesApp.getUi().alert(e.message)
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function testThumbs() {
  var slideIdArray = ['g730a5a125a_1_0', 'g72b54a5b54_0_22']
  
  viewModel.slideMeta.push( {id: 'g730a5a125a_1_0', thumbNail: 'http://somewhere.com/theimage.png1' })
  viewModel.slideMeta.push( {id: 'g72b54a5b54_0_22', thumbNail: 'http://somewhere.com/theimage2.png' })
  getThumbNailsForSlides(slideIdArray)
  
  for (const id of slideIdArray) {
    getThumbNail(id)
  } 
  
  var x = 1
}

// -------------------------------------------------------------------------------
//
// fetch from cached list
//
// -------------------------------------------------------------------------------
function getThumbNailsForSlides(slideIdArray) {
  var retObj = []
  if(slideIdArray != null && slideIdArray.length > 0) {  
    for(const slide of slideIdArray) {
      retObj.push(viewModel.slideMeta.find((x) => { return slide == x.id }))
    }
  }
  return retObj
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getThumbNails(slideIdArray) { 
  if(slideIdArray != null && slideIdArray.length > 0) {
    for(const slide of slideIdArray) {
      getThumbNail(slide.id)
    }
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getThumbNail(slideId) {

  //
  // Call Google Slide API async to get slide thumbnail
  //
  slideThumbTask = new Promise(function(completeTask, failTask){
    var parameters = {
      method: "GET",
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      contentType: "application/json",
      muteHttpExceptions: true
    };
    var httpUrl = "https://slides.googleapis.com/v1/presentations/" + SlidesApp.getActivePresentation().getId() + "/pages/" + slideId + "/thumbnail"
    var response = UrlFetchApp.fetch(httpUrl, parameters);
    completeTask({ id: slideId, urlData: JSON.parse(response.getContentText()) })
    
  }).then((a) => {
 
   //
   // work thro' the maintained list of slides and update
   // the thumbnail image for each one
   //
   viewModel.slideMeta = viewModel.slideMeta.map((x) => {
      if(a.id == x.id) {
            return {
              id: a.id,
              thumbNail: a.urlData.contentUrl,
              timestamp: null
       }
      } else {
        return x
      }
    })
  })
}

// -------------------------------------------------------------------------------
//
// //getElementsByClassName(o.slideThumbnails[0].id)[0].innerHTML = o.slideThumbnails[0]["url"];
//
// -------------------------------------------------------------------------------
function createSidebarHTML() {
  try {
    var currentPresentation = SlidesApp.getActivePresentation()
    var playLists = getData(app.PLAYLIST_RECORD_NAME)
    
    if(playLists != null && playLists.pl.length > 0) {
    
        var listString = '<html><head><link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">'
        listString = listString + '<script> function onSuccess(o) { document.getElementById("img-" + o.id).src = o.url; } function onLoadThumbnails(o) {  } ' 
        listString = listString + 'function toggle(elmId) { var elm = document.getElementById("list-" + elmId); elm.style.display = elm.style.display === "block" ? "none" : "block" } </script>' 
        listString = listString + '</head><body onload="google.script.run.withSuccessHandler(onLoadThumbnails).bobbins(\'' + playLists.pl + '\')" ><div class="sidebar"><div>Toolbar | Sort | Filter</div><hr/>'

        for (var i = 0; i < playLists.pl.length; i++) {
        var plName = playLists.pl[i].name 
        listString = listString + '<div class="block" ><span><h2>' + plName + '</h2><h3>' + playLists.pl[i].desc + '</h3>'
                                + '<button class="gray" style="cursor:hand;" onclick="google.script.run.deletePlayList(\'' + plName + '\')">Delete</button>'
                                + '<button class="blue" style="cursor:hand;" onclick="google.script.run.newFromPlaylist(\'' + plName + '\')">Export</button>'
                                + '<button class="gray" onclick="toggle(' + i + ')" style="cursor:hand;">+</button></span><div id="list-' + i + '" >'
        for (var j = 0; j < playLists.pl[i].ids.length; j++) {
          var thisSlideId = playLists.pl[i].ids[j]
          //
          // if slideId is missing from presentation then flag in list
          // also add to list of slides for which we'll need a thumbnail
          //
          var borderColor = '#ee1111'
          if(null != currentPresentation.getSlideById(thisSlideId)) {
            borderColor = '#eeeeee'
            //
            // Fetch or refresh our cached thumbnail for this slide
            //
            viewModel.updateThumbnailCache(thisSlideId)
            //listString = listString + '<button class="gray" style="cursor:hand;" onclick="google.script.run.withSuccessHandler(onSuccess).qwerty(\'' + thisSlideId + '\')">Update</button>'
            listString = listString + '<p><a style="color:back; text-decoration: none;" href="javascript:google.script.run.selectSlide(\'' + thisSlideId + '\')"><div class="' + thisSlideId + '"  style="cursor:hand; border-radius: 8px; background: #ffffff; border: 2px solid ' + borderColor + '; border-width:2px; color:black; padding: 2px;  width: 200px;  height: 110px;text-align: center;vertical-align: middle;"><div style="position:relative;top:0%;"><img style="height:110px;width:200px;" id="img-' + thisSlideId + '" src="' + getCachedUrl(thisSlideId) + '"/></div></div></a></p>';
            listString = listString + '<p>+ insert</p>'
          }
        }
        listString = listString + '</div></div><hr/>'
      }
      listString = listString + '</div></body></html>'
    }

    return listString
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }   
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function newFromPlaylist(name) {
  var currentPresentation = SlidesApp.getActivePresentation();
  var newName = currentPresentation.getName() + '-' + name
  var newPresentation = SlidesApp.create(newName)
  var pl = getPlaylistByName(name)
  if(pl != null) {
    for(var sid in pl[0].ids) {
      var slide = currentPresentation.getSlideById(pl[0].ids[sid])
      if(slide != null) {
        newPresentation.appendSlide(slide);
      }
      else {
        SlidesApp.getUi().alert('warning - slide missing ' + sid)
      }
    }
  }
  SlidesApp.getUi().alert('New presentation "' + newName + '" created')
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function onCreate() { 
  try {
  
    var playLists = getData(app.PLAYLIST_RECORD_NAME)
  
    var selectedRange = SlidesApp.getActivePresentation().getSelection()
    if(selectedRange == null) {
      SlidesApp.getUi().alert('You need to select more than one slide')
      return 
    }
        
    var selectedPageRange = selectedRange.getPageRange()
    if(selectedPageRange == null) {
      SlidesApp.getUi().alert('You need to select more than one slide')
      return 
    }
    
    var selectedslides = selectedPageRange.getPages()
    if(selectedslides == null || selectedslides.length < 2) {
      SlidesApp.getUi().alert('You need to slect more than one slide')
      return
    }
   
    var response = SlidesApp.getUi().prompt('Enter Playlist name:')
    var plName = ''
    var plDesc = ''
    if (response.getSelectedButton() == SlidesApp.getUi().Button.OK) {
      plName = response.getResponseText()
      
      if(plName == null || plName == '') {
        SlidesApp.getUi().alert('no name entered')
        return
      }
      
      for (const pl of playLists.pl) {
        if(plName == pl.name) {
          SlidesApp.getUi().alert('Playlist ' + plName + ' already exists, choose again')
          return 
        }
      }     
    }
    
    response = SlidesApp.getUi().prompt('Enter Playlist description:')
    if (response.getSelectedButton() == SlidesApp.getUi().Button.OK) {
      plDesc = response.getResponseText()   
    }    
    
    var playListIds = []
    for (var i = 0; i < selectedslides.length; i++) {
      var id = selectedslides[i].getObjectId()
      playListIds.push(id)
    }
    
    var playListEntry = {
      desc: plDesc,
      name: plName,
      ids: playListIds
    }
    playLists.pl.push(playListEntry)
    setData(app.PLAYLIST_RECORD_NAME, playLists)
    showSideBar();
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getPlaylistByName(name) {
  try {
        var playLists = getData(app.PLAYLIST_RECORD_NAME)
        var thePlaylist = playLists.pl.filter(function(value, index, arr){ return value.name == name})
        return thePlaylist
  } 
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
  return null
}


function testProps()
{
  var scriptProperties = PropertiesService.getScriptProperties()
  var pl = {"pl":[{"desc":"","name":"HollySmowly","ids":["g72a6e3d964_0_1","g72a6e3d964_0_5","g72a6e3d964_0_13","p","g72a6e3d964_0_9","g72b54a5b54_0_0","g72b54a5b54_0_7","g72b54a5b54_0_14","g72b54a5b54_0_18","g72b54a5b54_0_22","g730a5a125a_1_0"]},{"desc":"just the green slides","name":"Green deck","ids":["g72b54a5b54_0_18","g72b54a5b54_0_0","g72a6e3d964_0_5"]}]}
  var userProperties = PropertiesService.getUserProperties()
  //var data = { [app.PLAYLIST_RECORD_NAME]: pl }
  //var o = JSON.stringify(data)
  userProperties.setProperties({[app.PLAYLIST_RECORD_NAME]: JSON.stringify(pl)})
  
  var units = userProperties.getProperty(app.PLAYLIST_RECORD_NAME);
  o = JSON.parse(units)
}

//-------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function initDataStore() {
  try {
    if(getData(app.PLAYLIST_RECORD_NAME) == null) {
      setData(app.PLAYLIST_RECORD_NAME, app.DEFAULT_DATA_MODEL )
    }
  }
  catch (e) {
  }
}

//-------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function parkDataStore() {
  try {
    var pdr = getData(app.PLAYLIST_RECORD_NAME)
    if(pdr != null) {
      setData(app.PLAYLIST_RECORD_NAME_PARK, pdr )
      delData(app.PLAYLIST_RECORD_NAME_PARK)
      initDataStore() 
    }
  }
  catch (e) {
    //
  }
}

//-------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function restoreDataStore() {
  try {
    var pdr = getData(app.PLAYLIST_RECORD_NAME_PARK)
    if(pdr != null) {
      setData(app.PLAYLIST_RECORD_NAME, pdr )
      delData(app.PLAYLIST_RECORD_NAME_PARK)
    }
  }
  catch (e) {
    //
  }
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function delData(objectName, jsonObject) {
  var userProperties = PropertiesService.getUserProperties()
  userProperties.deleteProperty(objectName)
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function setData(objectName, jsonObject) {
  var userProperties = PropertiesService.getUserProperties()
  userProperties.setProperties({ [objectName]: JSON.stringify(jsonObject) })
}

function setDataString(objectName, stringValue) {
  var userProperties = PropertiesService.getUserProperties()
  userProperties.setProperties({ [objectName]: stringValue })
}

function getDataString(objectName) {
  var userProperties = PropertiesService.getUserProperties()
  return userProperties.getProperty(objectName)
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getData(objectName) {
  var userProperties = PropertiesService.getUserProperties()
  return JSON.parse(userProperties.getProperty(objectName))
}