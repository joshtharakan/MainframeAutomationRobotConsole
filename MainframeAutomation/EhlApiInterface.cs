using System;
interface EhlApiInterface
{
    //Boolean GetConnection();
    //void endConnection();
    //void PauseAndReset();
    //void StartHostNotification();
    //Boolean SendKey(String keyValue);
    //Boolean SetSessionParams();
    //String QuerySessionStatus();
    //short SearchField(short searchPosition, String searchData);
    //short SearchPS(short searchPosition, String searchData);
    //short FindFieldPosition(String searchData);
    //short FindFieldLength(String searchData);
    //Boolean CopyStringToField(short position, String data);
    //Boolean CopyStringToPS(short position, short length, String data);
    //Boolean ReserveSession();
    //Boolean ReleaseSession();
    //short SetCursor(short position);
    //Boolean InitialCheck(String expectedMenu);
    //short SearchAndPlaceCursor(String searchField, short offset);
    Boolean ExecuteCommand(String command, String expectedMenu);
    Boolean SearchFieldAndPopulate(String fieldName, String fieldValue);
    string GetStatus();
    //void CheckException(short returnCode);
    //Boolean DisconnectAndConnect();
    
}
