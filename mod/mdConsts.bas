Attribute VB_Name = "mdConsts"
'INTERNET_FLAG_IGNORE_CERT_CN_INVALID Disables Win32 Internet function checking of SSL/PCT-based certificates that are returned from the server against the host name given in the request. Win32 Internet functions use a simple check against certificates by comparing for matching host names and simple wildcarding rules.
'INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Disables Win32 Internet function checking of SSL/PCT-based certificates for proper validity dates.
'INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP Disables the ability of the Win32 Internet functions to detect this special type of redirect. When this flag is used, Win32 Internet functions transparently allow redirects from HTTPS to HTTP URLs.
'INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Disables the ability of the Win32 Internet functions to detect this special type of redirect. When this flag is used, Win32 Internet functions transparently allow redirects from HTTP to HTTPS URLs.
'INTERNET_FLAG_KEEP_CONNECTION Uses keep-alive semantics, if available, for the connection. This flag is required for Microsoft Network (MSN), NT LAN Manager (NTLM), and other types of authentication.
'INTERNET_FLAG_MAKE_PERSISTENT Not supported.
'INTERNET_FLAG_MUST_CACHE_REQUEST Causes a temporary file to be created if the file cannot be cached. Identical to the preferred value INTERNET_FLAG_NEED_FILE.
'INTERNET_FLAG_NEED_FILE Causes a temporary file to be created if the file cannot be cached.
'INTERNET_FLAG_NO_AUTH Does not attempt authentication automatically.
'INTERNET_FLAG_NO_AUTO_REDIRECT Does not automatically handle redirection in HttpSendRequest.
'INTERNET_FLAG_NO_CACHE_WRITE Does not add the returned entity to the cache.
'INTERNET_FLAG_NO_COOKIES Does not automatically add cookie headers to requests, and does not automatically add returned cookies to the cookie database.
'INTERNET_FLAG_NO_UI Disables the cookie dialog box.
'INTERNET_FLAG_PRAGMA_NOCACHE Forces the request to be resolved by the origin server, even if a cached copy exists on the proxy.
'INTERNET_FLAG_READ_PREFETCH This flag is currently disabled.
'INTERNET_FLAG_RELOAD Not supported. Data is always from the wire and caching is not supported.
'INTERNET_FLAG_RESYNCHRONIZE Reloads HTTP resources if the resource has been modified since the last time it was downloaded.
'INTERNET_FLAG_SECURE Uses SSL/PCT transaction semantics.
'HTTP_QUERY_ACCEPT  Retrieves the acceptable media types for the response.
'HTTP_QUERY_ACCEPT_CHARSET  Retrieves the acceptable character sets for the response.
'HTTP_QUERY_ACCEPT_ENCODING  Retrieves the acceptable content-coding values for the response.
'HTTP_QUERY_ACCEPT_LANGUAGE  Retrieves the acceptable natural languages for the response.
'HTTP_QUERY_ACCEPT_RANGES  Retrieves the types of range requests that are accepted for a resource.
'HTTP_QUERY_AGE  Retrieves the Age response-header field, which contains the sender's estimate of the amount of time since the response was generated at the origin server.
'HTTP_QUERY_ALLOW  Receives the methods supported by the server.
'HTTP_QUERY_AUTHORIZATION  Retrieves the authorization credentials used for a request.
'HTTP_QUERY_CACHE_CONTROL  Retrieves the cache control directives.
'HTTP_QUERY_CONNECTION  Retrieves any options that are specified for a particular connection and must not be communicated by proxies over further connections.
'HTTP_QUERY_COOKIE  Retrieves any cookies associated with the request.
'HTTP_QUERY_CONTENT_BASE  Retrieves the base URI for resolving relative URLs within the entity.
'HTTP_QUERY_CONTENT_DESCRIPTION  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_CONTENT_DISPOSITION  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_CONTENT_ENCODING  Receives any additional content codings that have been applied to the entire resource.
'HTTP_QUERY_CONTENT_ID  Receives the content identification.
'HTTP_QUERY_CONTENT_LANGUAGE  Receives the language that the content is in.
'HTTP_QUERY_CONTENT_LENGTH  Receives the size of the resource, in bytes.
'HTTP_QUERY_CONTENT_LOCATION Retrieves the resource location for the entity enclosed in the message.
'HTTP_QUERY_CONTENT_MD5  Retrieves a MD5 digest of the entity-body for the purpose of providing an end-to-end message integrity check (MIC) for the entity-body.
'HTTP_QUERY_CONTENT_RANGE Retrieves the location in the full entity-body where the partial entity-body should be inserted and the total size of the full entity-body.
'HTTP_QUERY_CONTENT_TRANSFER_ENCODING  Receives the additional content coding that has been applied to the resource.
'HTTP_QUERY_CONTENT_TYPE  Receives the content type of the resource (such as text/html).
'HTTP_QUERY_COST  No longer implemented.
'HTTP_QUERY_DATE  Receives the date and time at which the message was originated.
'HTTP_QUERY_DERIVED_FROM  No longer supported.
'HTTP_QUERY_ETAG  Retrieves the entity tag for the associated entity.
'HTTP_QUERY_EXPIRES  Receives the date and time after which the resource should be considered outdated.
'HTTP_QUERY_FORWARDED  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_FROM  Retrieves the e-mail address for the human user who controls the requesting user agent if the From header is given.
'HTTP_QUERY_HOST  Retrieves the Internet host and port number of the resource being requested.
'HTTP_QUERY_IF_MATCH  Retrieves the contents of the If-Match request-header field.
'HTTP_QUERY_IF_MODIFIED_SINCE  Retrieves the contents of the If-Modified-Since header.
'HTTP_QUERY_IF_NONE_MATCH  Retrieves the contents of the If-None-Match request-header field.
'HTTP_QUERY_IF_RANGE  Retrieves the contents of the If-Range request-header field. This header allows the client application to check if the entity related to a partial copy of the entity in the client application's cache has not been updated. If the entity has not been updated, send the parts that the client application is missing. If the entity has been updated, send the entire updated entity.
'HTTP_QUERY_IF_UNMODIFIED_SINCE  Retrieves the contents of the If-Unmodified-Since request-header field.
'HTTP_QUERY_LINK  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_LAST_MODIFIED  Receives the date and time at which the server believes the resource was last modified.
'HTTP_QUERY_LOCATION  Retrieves the absolute URI used in a Location response-header.
'HTTP_QUERY_MAX  Retrieves the maximum value of an HTTP_QUERY_* value.
'HTTP_QUERY_MAX_FORWARDS  Retrieves the number of proxies or gateways that can forward the request to the next inbound server.
'HTTP_QUERY_MESSAGE_ID  No longer implemented.
'HTTP_QUERY_MIME_VERSION  Receives the version of the MIME protocol that was used to construct the message.
'HTTP_QUERY_ORIG_URI  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_PRAGMA  Receives the implementation-specific directives that may apply to any recipient along the request/response chain.
'HTTP_QUERY_PROXY_AUTHENTICATE  Retrieves the authentication scheme and realm returned by the proxy.
'HTTP_QUERY_PROXY_AUTHORIZATION  Retrieves the header that is used to identify the user to a proxy that requires authentication.
'HTTP_QUERY_PUBLIC  Receives methods available at this server.
'HTTP_QUERY_RANGE  Retrieves the byte range of an entity.
'HTTP_QUERY_RAW_HEADERS  Receives all the headers returned by the server. Each header is terminated by "\0". An additional "\0" terminates the list of headers.
'HTTP_QUERY_RAW_HEADERS_CRLF  Receives all the headers returned by the server. Each header is separated by a carriage return/line feed (CR/LF) sequence.
'HTTP_QUERY_REFERER  Receives the URI of the resource where the requested URI was obtained.
'HTTP_QUERY_REFRESH  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_REQUEST_METHOD  Receives the verb that is being used in the request, typically GET or POST.
'HTTP_QUERY_RETRY_AFTER  Retrieves the amount of time the service is expected to be unavailable.
'HTTP_QUERY_SERVER  Retrieves information about the software used by the origin server to handle the request.
'HTTP_QUERY_SET_COOKIE  Receives the value of the cookie set for the request.
'HTTP_QUERY_STATUS_CODE  Receives the status code returned by the server.
'HTTP_QUERY_STATUS_TEXT  Receives any additional text returned by the server on the response line.
'HTTP_QUERY_TITLE  Obsolete. Maintained for legacy application compatibility only.
'HTTP_QUERY_TRANSFER_ENCODING  Retrieves the type of transformation that has been applied to the message body so it can be safely transferred between the sender and recipient.
'HTTP_QUERY_UPGRADE  Retrieves the additional communication protocols that are supported by the server.
'HTTP_QUERY_URI  Receives some or all of the Uniform Resource Identifiers (URIs) by which the Request-URI resource can be identified.
'HTTP_QUERY_USER_AGENT  Retrieves information about the user agent that made the request.
'HTTP_QUERY_VARY  Retrieves the header that indicates that the entity was selected from a number of available representations of the response using server-driven negotiation.
'HTTP_QUERY_VERSION  Receives the last response code returned by the server.
'HTTP_QUERY_VIA  Retrieves the intermediate protocols and recipients between the user agent and the server on requests, and between the origin server and the client on responses.
'HTTP_QUERY_WARNING  Retrieves additional information about the status of a response that may not be reflected by the response status code.
'HTTP_QUERY_WWW_AUTHENTICATE  Retrieves the authentication scheme and realm returned by the server.
'HTTP_QUERY_CUSTOM  Causes HttpQueryInfo to search for the header name specified in lpvBuffer and store the header information in lpvBuffer.
'HTTP_QUERY_FLAG_COALESCE  Not supported.
'HTTP_QUERY_FLAG_NUMBER  Returns the data as a 32-bit number for headers whose value is a number, such as the status code.
'HTTP_QUERY_FLAG_REQUEST_HEADERS  Queries request headers only.
'HTTP_QUERY_FLAG_SYSTEMTIME  Returns the header value as a standard Win32 SYSTEMTIME structure, which does not require the application to parse the data. Use for headers whose value is a date/time string, such as "Last-Modified-Time".
'INTERNET_AUTODIAL_FAILIFSECURITYCHECK
'Causes InternetAutodial to fail if file and printer sharing is disabled for Microsoft® Windows® 95 or later.
'INTERNET_AUTODIAL_FORCE_ONLINE
'Forces an online Internet connection.
'INTERNET_AUTODIAL_FORCE_UNATTENDED
'Forces an unattended Internet dial-up.
'FOR InternetCanonicalizeUrl
'ICU_BROWSER_MODE Does not encode or decode characters after "#" or "?", and does not remove trailing white space after "?". If this value is not specified, the entire URL is encoded, and trailing white space is removed.
'ICU_DECODE Converts all %XX sequences to characters, including escape sequences, before the URL is parsed.
'ICU_ENCODE_SPACES_ONLY Encodes spaces only.
'ICU_NO_ENCODE Does not convert unsafe characters to escape sequences.
'ICU_NO_META Does not remove meta sequences (such as "." and "..") from the URL.
'If no flags are specified (dwFlags = 0), the function converts all unsafe characters and meta sequences (such as \.,\ .., and \...) to escape sequences.
'ERROR_BAD_PATHNAME The URL could not be canonicalized.
'ERROR_INSUFFICIENT_BUFFER Canonicalized URL is too large to fit in the buffer provided. The *lpdwBufferLength parameter is set to the size, in bytes, of the buffer required to hold the resultant, canonicalized URL.
'ERROR_INTERNET_INVALID_URL The format of the URL is invalid.
'ERROR_INVALID_PARAMETER Bad string, buffer, buffer size, or flags parameter.
'end of InternetCanonicalizeUrl flags
'FLAG_ICC_FORCE_CONNECTION
'ICU_BROWSER_MODE Does not encode or decode characters after "#" or "?", and does not remove trailing white space after "?". If this value is not specified, the entire URL is encoded and trailing white space is removed.
'ICU_DECODE Converts all %XX sequences to characters, including escape sequences, before the URL is parsed.
'ICU_ENCODE_SPACES_ONLY Encodes spaces only.
'ICU_NO_ENCODE Does not convert unsafe characters to escape sequences.
'ICU_NO_META Does not remove meta sequences (such as "." and "..") from the URL.
'INTERNET_DEFAULT_FTP_PORT Uses the default port for FTP servers (port 21).
'INTERNET_DEFAULT_HTTP_PORT Uses the default port for HTTP servers (port 80).
'INTERNET_DEFAULT_HTTPS_PORT Uses the default port for HTTPS servers (port 443).
'INTERNET_DEFAULT_SOCKS_PORT Uses the default port for SOCKS firewall servers (port 1080).
'INTERNET_INVALID_PORT_NUMBER Uses the default port for the service specified by dwService.
'INTERNET_SERVICE_FTP FTP service.
'INTERNET_SERVICE_HTTP HTTP service.
'ICU_DECODE Converts encoded characters back to their normal form. This can be used only if the user provides buffers in the URL_COMPONENTS structure to copy the components into.
'ICU_ESCAPE Converts all escape sequences (%xx) to their corresponding characters. This can be used only if the user provides buffers in the URL_COMPONENTS structure to copy the components into.
'ICU_ESCAPE Converts all escape sequences (%xx) to their corresponding characters.
'ICU_USERNAME When adding the user name, uses the name that was specified at logon time.
'INTERNET_AUTODIAL_FORCE_ONLINE
'Forces an online connection.
'INTERNET_AUTODIAL_FORCE_UNATTENDED
'Forces an unattended Internet dial-up. If user intervention is required, the function will fail.
'INTERNET_DIAL_FORCE_PROMPT
'Ignores the "dial automatically" setting and forces the dialing user interface to be displayed.
'INTERNET_DIAL_UNATTENDED
'Connects to the Internet through a modem, without displaying a user interface, if possible. Otherwise, the function will wait for user input.
'INTERNET_DIAL_SHOW_OFFLINE
'Shows the Work Offline button instead of Cancel button in the dialing user interface.
'InternetErrorDlg
'ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR Notifies the user of the zone crossing to and from a secure site.
'ERROR_INTERNET_INCORRECT_PASSWORD Displays a dialog box for obtaining the user's name and password. (On Microsoft Windows® 95, the function first attempts to use any cached authentication information for the server being accessed, before displaying a dialog box.)
'ERROR_INTERNET_INVALID_CA Notifies the user that the Win32® Internet function does not recognize the certificate authority that generated the certificate for this Secure Sockets Layer (SSL) site.
'ERROR_INTERNET_POST_IS_NON_SECURE Displays a warning about posting data to the server through a nonsecure connection.
'ERROR_INTERNET_SEC_CERT_CN_INVALID Indicates that the SSL certificate Common Name (hostname field) is incorrect. Displays an Invalid SSL Common Name dialog box, and lets the user view the incorrect certificate. Also allows the user to select a certificate in response to a server request.
'ERROR_INTERNET_SEC_CERT_DATE_INVALID Tells the user that the SSL certificate has expired.
'dwFlags
'[in] Unsigned long integer value that contains the action flags. Can be a combination of these values:
'Value Description
'FLAGS_ERROR_UI_FILTER_FOR_ERRORS Scans the returned headers for errors. Call after using HttpSendRequest. This option detects any hidden errors, such as an authentication error.
'FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS If the function succeeds, stores the results of the dialog box in the Internet handle.
'FLAGS_ERROR_UI_FLAGS_GENERATE_DATA Queries the Internet handle for needed information. The function constructs the appropriate data structure for the error. (For example, for Cert CN failures, the function grabs the certificate.)
'FLAGS_ERROR_UI_SERIALIZE_DIALOGS Serializes authentication dialog boxes for concurrent requests on a password cache entry. The lppvData parameter should contain the address of a pointer to an INTERNET_AUTH_NOTIFY_DATA structure, and the client should implement a thread-safe, nonblocking callback function.
'end of InternetErrorDlg
'InternetSetFilePointer
'FILE_BEGIN Starting point is zero or the beginning of the file. If FILE_BEGIN is specified, lDistanceToMove is interpreted as an unsigned location for the new file pointer.
'FILE_CURRENT Current value of the file pointer is the starting point.
'FILE_END Current end-of-file position is the starting point. This method fails if the content length is unknown.
'end of InternetSetFilePointer
'INTERNET_CACHE_GROUP_ADD Adds the cache entry to the cache group.
'INTERNET_CACHE_GROUP_REMOVE Removes the cache entry from the cache group.
'CACHE_ENTRY_ACCTIME_FC
'Sets the last access time.
'CACHE_ENTRY_ATTRIBUTE_FC
'Sets the cache entry type.
'CACHE_ENTRY_EXEMPT_DELTA_FC
'Sets the exempt delta.
'CACHE_ENTRY_EXPTIME_FC
'Sets the expire time.
'CACHE_ENTRY_HEADERINFO_FC
'Not currently implemented.
'CACHE_ENTRY_HITRATE_FC
'Sets the hit rate.
'CACHE_ENTRY_MODTIME_FC
'Sets the last modified time.
'CACHE_ENTRY_SYNCTIME_FC
'Sets the last sync time.
