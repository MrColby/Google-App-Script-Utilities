/*
    This script is designed to help control in-box bloat. 

    If you are the sort of person who uses their in-box as long-term storage, don't use this script.

    The text of this script needs to be saved in google drive as a Google Apps Script. If you don't see that option in the New dropdown menu, then select +Connect More Apps.
    This script will start at a specified point down in your message threads, then delete up to the max number of messages specified. It can be set to delete only unread messages (I figure that if I haven't read it and another 1000 messages or more have come in on top if it, I'm never going to read it), or all messages.
    To set it to run automatically (every day, every hour, whatever you want), then set a trigger. Click the clock icon in the left-most panel of the app script editor window. Choose the deleteMessages function to run. Choose the Head deployment. Select the Time-driven option. I like the Day timer so mine runs once in the middle of the night. I've set the failure notification to notify me immediately. 
*/


// Set START_INDEX to any positive whole number to control how many most-recent threads to leave unaffected.
var START_INDEX = 1001;

// Set MESSAGES_TO_DELETE to any positive whole number to control how many messages to potentially delete. If the script times out and the error messages bother you, set MESSAGES_TO_DELETE to a lower value.
var MESSAGES_TO_DELETE = 1000;

// SET ONLY_UNREAD to control whether only unread messages will be deleted.
var ONLY_UNREAD = false;

// set how many threads to retrieve at a time. There's a performance question that I don't have the answer to. What's the best value? There's also the issue of the Gmail service quota of daily calls.
var THREAD_INCREMENT = 500;

function deleteMessages() {
  var messagesDeleted = 0;
  var threads = null;
  var deleting = true;
  while (deleting) {
    var threads = GmailApp.getInboxThreads(START_INDEX, THREAD_INCREMENT);
    for (var i=0; deleting && i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      for (var j=0; deleting && j < messages.length; j++) {
        var message = messages[j];
        if (ONLY_UNREAD && message.isUnread()) {
          message.moveToTrash();
          messagesDeleted++;
        }
        else {
          message.moveToTrash();
          messagesDeleted++;
        }
        if (messagesDeleted >= MESSAGES_TO_DELETE) {
          deleting = false;
        }

      }
    }
    if (messagesDeleted >= MESSAGES_TO_DELETE) {
      deleting = false;
    }
  }
  Logger.log(messagesDeleted + " messages deleted.");
  return messagesDeleted;
}
