# -EAP-MI

本程式負責處理MI機台PDS訊息，接收VFEI事件後依批次與批內Lot匹配槽位，組裝PDS回傳並計算參數差值，如AMU_DELTA、ENERGY_DELTA、ASCANS_DELTA，以及特定槽位ACCEL_I處理，最後送回Host並記錄Log以利追蹤。

本專案為MI機台EAP雙Task，涵蓋Lot/Batch管理、VFEI/Host介接、Recipe與PDS計算、任務調度與SMIF/RM搬運；提供UI查詢與警報，強化監控、差值運算、搬運追蹤，支援雙機同步與槽位邏輯。