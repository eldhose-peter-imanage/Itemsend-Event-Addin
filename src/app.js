var mailboxItem;
var itemSendEvent;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateBody(event) {
    itemSendEvent = event;
    console.log("Inside Validate body - before openDialog");
    openDialog();
    console.log("Inside Validate body - after openDialog");
    //call after openDialog is completed.
    //mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
}

function f1(){
    {
        console.log("Returned from dialog");
        console.log("mes : "+mes );
        console.log(event);
        console.log(itemSendEvent);

        if(mes === "tag"){
            //mailboxItem.notificationMessages.addAsync('Send', { type: 'informationalMessage', message: 'The user has clicked add tag button' });
            //change Subject.
            changeSubject();
            // Allow send.
         
        }
        else if(mes === "cancel"){
           // mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'The user has clicked cancel button' });
            // Block send.
            itemSendEvent.completed({ allowEvent: false });
        }
    }
}

let subject;

function changeSubject(){

    mailboxItem.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                console.log(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                subject = asyncResult.value + "-Taged";
                setSubject();
            }
        });
 
}

function setSubject(){
    console.log("Subject : "+ subject);
    mailboxItem.subject.setAsync(
        subject,
        { asyncContext: { itemSendEvent } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                console.log(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
                itemSendEvent.completed({ allowEvent: true });
            }
        });
}
