#pragma once

enum class AddInType
{
	XL,
	OFFICE,
	NONE
};
enum class ImageType
{
	x64,
	x86
};

class XLaddinInformation
{
public:
	std::wstring	Description_;
	std::wstring	Name_;
	std::wstring	ProgId_;
	std::wstring	Installed_;
	std::wstring	FullName_;
	DWORD			LoadBehavior_;
	AddInType		addType_;
	std::wstring	key_;
	HKEY			parent_;
	std::wstring	str_account;
};



class ProcessInformation
{
public:
	std::wstring									Name_;
	ImageType										imagetype_;
	std::list<std::wstring>							modules_;
	std::map<std::wstring, XLaddinInformation>		addininformation_;
};


class AddinInformation
{
public:
	std::wstring Progid_;
	std::wstring Description_;
	std::wstring Software_;
	DWORD		 Startmode_;
	std::wstring Connected_;
	std::wstring Key_;
	HKEY		 Parent_;
	std::wstring FriendlyName_;
	std::wstring str_account_;
};