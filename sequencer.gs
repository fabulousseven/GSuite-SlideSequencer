// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
var app = {
  PLAYLIST_RECORD_NAME: "playlists",
  PLAYLIST_TABLE_NAME: ""
}
 
// -------------------------------------------------------------------------------
//
// OnInstall: build the meui items and other relevant UI assets
//
// -------------------------------------------------------------------------------
function onInstall(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Create new playlist', 'onCreate')
      .addItem('View playlists', 'showSideBar')
      .addToUi();
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function showSideBar() { 

  try {
    var playListSideBarHTML = createSidebarHTML()
    var htmlOutput = HtmlService.createHtmlOutput(playListSideBarHTML).setTitle('Sequencer Playlists');
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
     var playLists = getDataObjectByName(app.PLAYLIST_RECORD_NAME, '')
     var updatedPlaylist = playLists.pl.filter(function(value, index, arr){ return value.name != name;});
     playLists.pl = updatedPlaylist
     saveDataObjectByName(playLists, app.PLAYLIST_RECORD_NAME, '')
     showSideBar();
  }
  catch (e) {
  SlidesApp.getUi().alert(e.message)
  }
}

function debug() {
  playFromPlaylist('hope')
}

// -------------------------------------------------------------------------------
//
// Send the slides in a playlistas email attachments
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
function createSidebarHTML() {
  try {
    var playLists = getDataObjectByName(app.PLAYLIST_RECORD_NAME, '')
    var listString = '<html><head><link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">'
        listString = listString + '<script></script>' 
        listString = listString + '</head><body><div class="sidebar"><div>Toolbar | Sort | Filter</div>'
    
    if(playLists != null && playLists.pl.length > 0) {
      for (var i = 0; i < playLists.pl.length; i++) {
        playLists.pl[i].name 
        listString = listString + '<div class="block" ><span><h2>' + playLists.pl[i].name + '</h2><h3>' + playLists.pl[i].desc + '</h3>'
                                + '<button class="gray" style="cursor:hand;" onclick="google.script.run.playFromPlaylist(\'' + playLists.pl[i].ids + '\')">Play</button>'
                                + '<button class="gray" style="cursor:hand;" onclick="google.script.run.deletePlayList(\'' + playLists.pl[i].name + '\')">Delete</button>'
                                + '<button class="blue" style="cursor:hand;" onclick="google.script.run.newFromPlaylist(\'' + playLists.pl[i].name + '\')">New From</button></span>'
        for (var j = 0; j < playLists.pl[i].ids.length; j++) {
          listString = listString + '<p><a style="color:back; text-decoration: none;" href="javascript:google.script.run.selectSlide(\'' + playLists.pl[i].ids[j] + '\')"><div style="cursor:hand; border-radius: 8px;  background: #ffffff; border: 2px solid #eeeeee; border-width:2px; color:black; padding: 10px;  width: 140px;  height: 70px;">' + playLists.pl[i].ids[j] + '</div></a></p>';
        }
        listString = listString + '</div><hr/>'
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
  
    var playLists = getDataObjectByName(app.PLAYLIST_RECORD_NAME, '')
  
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
    saveDataObjectByName(playLists, app.PLAYLIST_RECORD_NAME, app.PLAYLIST_TABLE_NAME)
    showSideBar();
  }
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function buildPlaylist() {
  var listString = ''
  for (var i = 0; i < 12; i++) {
    listString = listString + '<p><span>Playlist: ' + i + '</span><span><input onclick="javascript:alert(\'yes\');" value="Build Deck" type=button/></span></p>'
  }
  return listString
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function upsertPlaylist(name, slideList) {

  var table = getNamedDataObject(name, table)
  var selectedslides = SlidesApp.getActivePresentation().getSelection()
  if(table) {
      for (var i = 0; i < 12; ++i) {
        listString = listString + '<p><span>Playlist: ' + i + '</span><span><input onclick="javascript:alert(\'yes\');" value="Build Deck" type=button/></span></p>'
      }  
  } 
  else {
  }
  return listString
}

// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getPlaylistByName(name) {
  try {
        var playLists = getDataObjectByName(app.PLAYLIST_RECORD_NAME, '')
        var thePlaylist = playLists.pl.filter(function(value, index, arr){ return value.name == name;});
        return thePlaylist
  } 
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
  return null
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getDataObjectByName(name, table) {
  try {
    var dataTable = getDataTable(table)
    for(i = 0; i < dataTable.getNumRows() ; i++) {
      var keyRange = dataTable.getCell(i, 0).getText()
      var key = keyRange.getRange(0, keyRange.getEndIndex() - 1).asString() 
      if(key === name) {
        var textData = dataTable.getCell(i, 1).getText().asString()
        var jsonData = JSON.parse(textData)
        return jsonData
      }
    }
  } 
  catch (e) {
    SlidesApp.getUi().alert(e.message)
  }
  return null
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function saveDataObjectByName(jsonObject, name, table) {
  try {
    var dataTable = getDataTable(table)
    for(i = 0; i < dataTable.getNumRows() ; i++) {
      var keyRange = dataTable.getCell(i, 0).getText()
      var key = keyRange.getRange(0, keyRange.getEndIndex() - 1).asString() 
      if(key === name) {
        var jsonDataString = JSON.stringify(jsonObject)
        dataTable.getCell(i, 1).getText().setText(jsonDataString)
        return jsonDataString
      }
    }
  } 
  catch (e) {
    SlidesApp.getUi().alert('data table missing in last slide')
  }
  return null
}
// -------------------------------------------------------------------------------
//
// -------------------------------------------------------------------------------
function getDataTable(table) {
  try {
    var slides = SlidesApp.getActivePresentation().getSlides();
    var dataSlide = slides[slides.length-1]
    var dataTable = dataSlide.getPageElements()[0].asTable()
    return dataTable
  } 
  catch (e) {
    SlidesApp.getUi().alert('data table missing in last slide')
  }
}

