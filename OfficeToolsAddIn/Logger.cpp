#include "pch.h"
#include "Logger.h"
#include "XLSingleton.h"



BOOST_LOG_ATTRIBUTE_KEYWORD(line_id, "LineID", unsigned int)
BOOST_LOG_ATTRIBUTE_KEYWORD(timestamp, "TimeStamp", boost::posix_time::ptime)
BOOST_LOG_ATTRIBUTE_KEYWORD(severity, "Severity", logging::trivial::severity_level)
BOOST_LOG_ATTRIBUTE_KEYWORD(processid, "ProcessID", logging::attributes::current_process_id::value_type)


class customlog_sink : public sinks::basic_formatted_sink_backend<char, sinks::concurrent_feeding>
{
public:
	void consume(const logging::record_view& rec, const string_type& fstring)
	{
		std::ostringstream os;
		os << "<" << rec[boost::log::trivial::severity] << "> " << fstring;
		HWND hwnd = XLSingleton::getInstance()->Get_Log_info();
		if (hwnd != nullptr)
		{
			if (::IsWindow(hwnd))
			{
				::SendMessageA(hwnd, LB_ADDSTRING, NULL,(LPARAM) os.str().c_str());
			}
		}		
	}
	customlog_sink()
	{
	}


private:
	
};


HRESULT GetLog(fs::path& pFileName)
{
	std::wstring		pFileNamePath;
	std::wstring		pFileNameNoEx;
	TCHAR				buffer[MAX_PATH * sizeof(TCHAR)];
	LPWSTR				wszPath = NULL;
	HRESULT				hr = S_OK;

	hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_CREATE, NULL, &wszPath);
	if (SUCCEEDED(hr))
	{
		GetModuleFileName(NULL, buffer, MAX_PATH);
		std::wstring::size_type pos = std::wstring(buffer).find_last_of(_T("\\/"));
		std::wstring::size_type pos1 = std::wstring(buffer).find_last_of(_T("."));

		if (pos != 0 && pos1 != 0 && pos1 > pos)
		{
			pFileNameNoEx = std::wstring(buffer).substr(pos + 1, pos1 - pos - 1);
			pFileNamePath = std::wstring(wszPath) + _T("\\temp\\") + pFileNameNoEx;
			pFileName = pFileNamePath + _T("\\") + pFileNameNoEx;
			if (!fs::is_directory(pFileNamePath))
			{
				if (!fs::create_directory(pFileNamePath))
				{
					hr = HRESULT_FROM_WIN32(GetLastError());
				}
			}
		}
		else
		{
			hr = HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
		}
	}
	return (hr);
}

BOOST_LOG_GLOBAL_LOGGER_INIT(logger, src::severity_logger_mt) 
{
	src::severity_logger_mt<boost::log::trivial::severity_level>	logger;
	fs::path														pFileName;

	// add attributes
	logger.add_attribute("LineID", attrs::counter<unsigned int>(1));     // lines are sequentially numbered
	logger.add_attribute("TimeStamp", attrs::local_clock());             // each log line gets a timestamp
																	 
	typedef sinks::synchronous_sink<sinks::text_file_backend> TextSink; // add a text sink

	typedef sinks::synchronous_sink< customlog_sink > sink_t;
	boost::shared_ptr< sink_t > sink(new sink_t());
	boost::shared_ptr< logging::core > core = logging::core::get();
	core->add_sink(sink);

//#if _DEBUG
	typedef sinks::synchronous_sink<sinks::debug_output_backend> outputdebugstring_sink;
	boost::shared_ptr<outputdebugstring_sink> output_sink(new outputdebugstring_sink());
//#endif // _DEBUG
	if (GetLog(pFileName) == S_OK)
	{

		

		boost::shared_ptr<sinks::text_file_backend> backend1 = boost::make_shared<sinks::text_file_backend>(
			keywords::file_name = pFileName.generic_string() + "sign_%Y-%m-%d_%H-%M-%S.%N.log",
			keywords::rotation_size = 10 * 1024 * 1024,
			keywords::time_based_rotation = sinks::file::rotation_at_time_point(0, 0, 0),
			keywords::min_free_space = 30 * 1024 * 1024);

		backend1->auto_flush(true);	

		boost::shared_ptr<TextSink> sink(new TextSink(backend1));

		logging::formatter formatter = expr::stream
			<< std::setw(7) << std::setfill('0') << line_id << std::setfill(' ') << " | "
			<< expr::format_date_time(timestamp, "%Y-%m-%d, %H:%M:%S.%f") << " "
			<< "[" << logging::trivial::severity << "]"
			<< " - " << expr::smessage;
		sink->set_formatter(formatter);

		//#ifndef DEBUG // we tracing only error and fatal on release mode
		//		sink->set_filter(severity >= 4);
		//#endif // DEBUG
		

//#if _DEBUG
		logging::formatter formatter1 = expr::stream
			<< std::setw(7) << std::setfill('0') << line_id << std::setfill(' ') << " | "
			<< "[" << expr::format_date_time(
				timestamp, "%Y-%m-%d %H:%M:%S") << "]"
			<< "[" << severity << "] "
			<< expr::smessage
			<< std::endl;

		output_sink->set_formatter(formatter1);
		logging::core::get()->add_sink(output_sink);
//#endif // _DEBUG
		// "register" our sink
		logging::core::get()->add_sink(sink);		
	}
	return logger;
}

