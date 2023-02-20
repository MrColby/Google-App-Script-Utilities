/*
    This script is designed to help control in-box bloat. 

    If you are the sort of person who uses their in-box as long-term storage, don't use this script.

    The text of this script needs to be saved in google drive as a Google Apps Script. If you don't see that option in the New dropdown menu, then select +Connect More Apps.
    This script will start at a specified point down in your message threads, then delete up to the max number of threads specified. It can be set to delete only unread threads (I figure that if I haven't read it and another 1000 threads or more have come in on top if it, I'm never going to read it), or all threads.
    To set it to run automatically (every day, every hour, whatever you want), then set a trigger. Click the clock icon in the left-most panel of the app script editor window. Choose the deleteThreads function to run. Choose the Head deployment. Select the Time-driven option. I like the Day timer so mine runs once in the middle of the night. I've set the failure notification to notify me immediately. 
*/


// Set START_INDEX to any positive whole number to control how many most-recent threads to leave unaffected.
var START_INDEX = 1001;

// Set THREADS_TO_DELETE to any positive whole number to control how many threads to potentially delete. If the script times out and the error messages bother you, set THREADS_TO_DELETE to a lower value.
var THREADS_TO_DELETE = 500;

// SET ONLY_READ to control whether only read threads will be deleted.
var ONLY_READ = false;



function deleteThreads() {
  var threadsDeleted = 0;
  var threads = GmailApp.getInboxThreads(START_INDEX, THREADS_TO_DELETE);
  for (var i=0; i < threads.length; i++) {
    var thread = threads[i];
    if (ONLY_READ && !thread.isUnread()) {
      thread.moveToTrash();
      threadsDeleted++;
    }
    else {
      thread.moveToTrash();
      threadsDeleted++;
    }
    if (threadsDeleted >= THREADS_TO_DELETE) {
      break;
    }
  }
  Logger.log(threadsDeleted + " threads deleted.");
  return threadsDeleted;
}
