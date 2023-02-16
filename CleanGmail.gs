/*
    This script is designed to help control in-box bloat. 
    If you are the sort of person who uses their in-box as long-term storage, don't use this script. But really?

    The text of this script needs to be save in google drive as a Google Apps Script. If you don't see that option in the New dropdown menu, then select +Connect More Apps.

    This script will start at a specified point down in your message threads, then delete up to the max number of threads supported by gmail (500 at the time I wrote this). It can be set to delete only unread messages (I figure that if I haven't read it and another 1000 messages or more have come in on top if it, I'm never going to read it), or all messages.

    To set it to run automatically (every day, every hour, whatever you want), then set a trigger. Click the clock icon in the left-most panel of the app script editor window. Choose the deleteThreads function to run. Choose the Head deployment. Select the Time-driven option. I like the Day timer so mine runs once in the middle of the night. I've set the failure notification to notify me daily. 
*/


// Set START_INDEX to any positive whole number to control how many most-recent threads to leave unaffected.
var START_INDEX = 1001;

// Set THREADS to any positive whole number <= 500 to control how many threads to potentially delete. If the script times out and the error messages bother you, set THREADS to a lower value.
var THREADS = 500;

// SET ONLY_UNREAD to control whether only unread messages will be deleted.
var ONLY_UNREAD = false;

function deleteThreads() {
  Logger.log("Inbox unread count before executions: " + GmailApp.getInboxUnreadCount());
  var threads = GmailApp.getInboxThreads(START_INDEX, THREADS);
  Logger.log(threads.length + " before run.");
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    for (var j=0; j < messages.length; j++) {
      var message = messages[j];
      if (ONLY_UNREAD && message.isUnread()) {
        message.moveToTrash();
      }
      else {
        message.moveToTrash();
      }
    }
  }
  Logger.log("Inbox unread count after executions: " + GmailApp.getInboxUnreadCount());
}
