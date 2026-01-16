# -EAP-MI

本程式負責處理MI機台PDS訊息，接收VFEI事件後依批次與批內Lot匹配槽位，組裝PDS回傳並計算參數差值，如AMU_DELTA、ENERGY_DELTA、ASCANS_DELTA，以及特定槽位ACCEL_I處理，最後送回Host並記錄Log以利追蹤。