
import { readExcel, readUseExcel, writeExcel, writeSelectedRange, writeNonExcel, openMessageBox, openMyDialog } from "@/utils/excel";
import useDialog from '@/hooks/useDialog'; // Import the custom hook

Office.onReady(() => {
});

const url = "https://localhost:3000/popup.html", height = 45, width = 55


function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}


function isValidIdentityCard(id) {
  // 验证身份证号是否为18位数字
  return /^[0-9]{18}$/.test(id);
}

function extractBirthday(id) {
  // 提取出生年月
  return id.substring(6, 14);
}

function openPopup(event) {

  // writeExcel("C1", 2)
  let dialog = null
  Office.context.ui.displayDialogAsync(
    url,
    { height, width },
    (result) => {
      dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        dialog.close();
      });
    }
  );
}


function btnConnectService(event) {
  
}

// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("openPopup", openPopup);
Office.actions.associate("btnConnectService", btnConnectService);
