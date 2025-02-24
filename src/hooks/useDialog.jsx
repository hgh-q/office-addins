import { useState } from 'react';

const useDialog = () => {
  const [dialog, setDialog] = useState(null);

  const openDialog = (url, processMessage, height = 45, width = 55) => {
    Office.context.ui.displayDialogAsync(
      url,
      { height, width },
      (result) => {
        const dialogInstance = result.value;
        setDialog(dialogInstance);
        dialogInstance.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    );
  };

  const closeDialog = () => {
    closeDialog()
    if (dialog) {
      dialog.close();
    }
  };

  return { openDialog, closeDialog, dialog };
};

export default useDialog;
