#pragma once

#include "ExcelProcessInformation.h"



class XLSingleton 
{
public:
  
    static class XLSingleton *instance;

    ProcessInformation	processinformation_;
    HWND                wndOutputDebug_;

   // Private constructor so that no objects can be created.
   class XLSingleton()
   {
       wndOutputDebug_ = nullptr;
   }

   public:

       XLSingleton::~XLSingleton()
       {
           int x = 0;
       }


   static XLSingleton *getInstance() 
   {
       if (!instance)
           instance = new XLSingleton;
      return instance;
   }

   ProcessInformation Get_Addin_info()
   {
      return this->processinformation_;
   }

   void Set_Addin_info(ProcessInformation pInfo)
   {
       processinformation_ = pInfo;
   }

   void Set_Log_info(HWND wndOutputDebug)
   {
       wndOutputDebug_ = wndOutputDebug;
   }
   HWND Get_Log_info()
   {
       return (wndOutputDebug_);
   }
};