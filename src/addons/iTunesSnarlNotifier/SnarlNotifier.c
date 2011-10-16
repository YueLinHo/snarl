/*
	File:		iTunesVisualSample.c

    Contains:   iTunes Visual Plug-ins sample code
 
    Version:    Technology: iTunes
                Release:    4.1

	Copyright: 	© Copyright 2003 Apple Computer, Inc. All rights reserved.
	
	Disclaimer:	IMPORTANT:  This Apple software is supplied to you by Apple Computer, Inc.
				("Apple") in consideration of your agreement to the following terms, and your
				use, installation, modification or redistribution of this Apple software
				constitutes acceptance of these terms.  If you do not agree with these terms,
				please do not use, install, modify or redistribute this Apple software.

				In consideration of your agreement to abide by the following terms, and subject
				to these terms, Apple grants you a personal, non-exclusive license, under Apple’s
				copyrights in this original Apple software (the "Apple Software"), to use,
				reproduce, modify and redistribute the Apple Software, with or without
				modifications, in source and/or binary forms; provided that if you redistribute
				the Apple Software in its entirety and without modifications, you must retain
				this notice and the following text and disclaimers in all such redistributions of
				the Apple Software.  Neither the name, trademarks, service marks or logos of
				Apple Computer, Inc. may be used to endorse or promote products derived from the
				Apple Software without specific prior written permission from Apple.  Except as
				expressly stated in this notice, no other rights or licenses, express or implied,
				are granted by Apple herein, including but not limited to any patent rights that
				may be infringed by your derivative works or by other works in which the Apple
				Software may be incorporated.

				The Apple Software is provided by Apple on an "AS IS" basis.  APPLE MAKES NO
				WARRANTIES, EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION THE IMPLIED
				WARRANTIES OF NON-INFRINGEMENT, MERCHANTABILITY AND FITNESS FOR A PARTICULAR
				PURPOSE, REGARDING THE APPLE SOFTWARE OR ITS USE AND OPERATION ALONE OR IN
				COMBINATION WITH YOUR PRODUCTS.

				IN NO EVENT SHALL APPLE BE LIABLE FOR ANY SPECIAL, INDIRECT, INCIDENTAL OR
				CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE
				GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
				ARISING IN ANY WAY OUT OF THE USE, REPRODUCTION, MODIFICATION AND/OR DISTRIBUTION
				OF THE APPLE SOFTWARE, HOWEVER CAUSED AND WHETHER UNDER THEORY OF CONTRACT, TORT
				(INCLUDING NEGLIGENCE), STRICT LIABILITY OR OTHERWISE, EVEN IF APPLE HAS BEEN
				ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

     Bugs?:      For bug reports, consult the following page on
                 the World Wide Web:
 
                     http://developer.apple.com/bugreporter/

*/

	// Windows headers
#include <windows.h>
#include <stdio.h> 
#include <tchar.h>

#include "iTunesVisualAPI.h"
#include "Snarl.h"

	#define GRAPHICS_DEVICE	HWND

#define	MAIN iTunesPluginMain
#define IMPEXP	__declspec(dllexport)

#define kSampleVisualPluginName		"SnarlNotifier"
#define	kSampleVisualPluginCreator	'FFP_'

#define	kSampleVisualPluginMajorVersion		1
#define	kSampleVisualPluginMinorVersion		0
#define	kSampleVisualPluginReleaseStage		0x80
#define	kSampleVisualPluginNonFinalRelease	0

enum
{
	kSettingsDialogResID	= 1000,
	
	kSettingsDialogOKButton	= 1,
	kSettingsDialogCancelButton,
	
	kSettingsDialogCheckBox1,
	kSettingsDialogCheckBox2,
	kSettingsDialogCheckBox3
};

struct VisualPluginData {
	void *				appCookie;
	ITAppProcPtr		appProc;
	
	GRAPHICS_DEVICE		destPort;
	Rect				destRect;
	OptionBits			destOptions;
	UInt32				destBitDepth;

	RenderVisualData	renderData;
	UInt32				renderTimeStampID;
	
	SInt8				waveformData[kVisualMaxDataChannels][kVisualNumWaveformEntries];
	
	UInt8				level[kVisualMaxDataChannels];		/* 0-128 */
	
	ITTrackInfoV1		trackInfo;
	ITStreamInfoV1		streamInfo;

	Boolean				playing;
	Boolean				padding[3];

/*
	Plugin-specific data
*/
};
typedef struct VisualPluginData VisualPluginData;

// ClearMemory
//
static void ClearMemory (LogicalAddress dest, SInt32 length)
{
	register unsigned char	*ptr;

	ptr = (unsigned char *) dest;
	
	if( length > 16 )
	{
		register unsigned long	*longPtr;
		
		while( ((unsigned long) ptr & 3) != 0 )
		{
			*ptr++ = 0;
			--length;
		}
		
		longPtr = (unsigned long *) ptr;
		
		while( length >= 4 )
		{
			*longPtr++ 	= 0;
			length		-= 4;
		}
		
		ptr = (unsigned char *) longPtr;
	}
	
	while( --length >= 0 )
	{
		*ptr++ = 0;
	}
}

/*
	ProcessRenderData
*/
static void ProcessRenderData (VisualPluginData *visualPluginData, const RenderVisualData *renderData)
{
	SInt16		index;
	SInt32		channel;
	
	visualPluginData->level[0] = 0;
	visualPluginData->level[1] = 0;
		
	if (renderData == nil)
	{
		ClearMemory(&visualPluginData->renderData, sizeof(visualPluginData->renderData));
		return;
	}

	visualPluginData->renderData = *renderData;
	
	for (channel = 0; channel < kVisualMaxDataChannels; channel++)
	{
		for (index = 0; index < kVisualNumWaveformEntries; index++)
		{
			SInt8		value;
			
			value = renderData->waveformData[channel][index] - 0x80;
			
			visualPluginData->waveformData[channel][index] = value;
				
			if (value < 0)
				value = -value;
				
			if (value > visualPluginData->level[channel])
				visualPluginData->level[channel] = value;
		}
	}
}


/*
	RenderVisualPort
*/
static void RenderVisualPort (VisualPluginData *visualPluginData, GRAPHICS_DEVICE destPort, const Rect *destRect, Boolean onlyUpdate)
{

	(void) visualPluginData;
	(void) onlyUpdate;
	
	if (destPort == nil)
		return;

	{
		/*
		RECT	srcRect;
		HBRUSH	hBrush;
		HDC		hdc;
		
		srcRect.left = destRect->left;
		srcRect.top = destRect->top;
		srcRect.right = destRect->right;
		srcRect.bottom = destRect->bottom;
		
		hdc = GetDC(destPort);		
		hBrush = CreateSolidBrush(RGB((UInt16)visualPluginData->level[1]<<1, (UInt16)visualPluginData->level[1]<<1, (UInt16)visualPluginData->level[0]<<1));
		FillRect(hdc, &srcRect, hBrush);
		DeleteObject(hBrush);
		ReleaseDC(destPort, hdc);
		*/
	}
}

/*
	AllocateVisualData is where you should allocate any information that depends
	on the port or rect changing (like offscreen GWorlds).
*/
static OSStatus AllocateVisualData (VisualPluginData *visualPluginData, const Rect *destRect)
{
	OSStatus		status;
	(void) visualPluginData;
	(void) destRect;

	status = noErr;
	return status;
}


/*
	DeallocateVisualData is where you should deallocate the things you have allocated
*/
static void DeallocateVisualData (VisualPluginData *visualPluginData)
{
		(void)visualPluginData;
}

static Boolean RectanglesEqual(const Rect *r1, const Rect *r2)
{
	if (
		(r1->left == r2->left) &&
		(r1->top == r2->top) &&
		(r1->right == r2->right) &&
		(r1->bottom == r2->bottom)
		)
		return true;
	return false;
	
}

// ChangeVisualPort
//
static OSStatus ChangeVisualPort (VisualPluginData *visualPluginData, GRAPHICS_DEVICE destPort, const Rect *destRect)
{
	OSStatus		status;
	Boolean			doAllocate;
	Boolean			doDeallocate;
	
	status = noErr;
	
	doAllocate		= false;
	doDeallocate	= false;
		
	if (destPort != nil)
	{
		if (visualPluginData->destPort != nil)
		{
			if (RectanglesEqual(destRect, &visualPluginData->destRect) == false)
			{
				doDeallocate	= true;
				doAllocate		= true;
			}
		}
		else
		{
			doAllocate = true;
		}
	}
	else
	{
		doDeallocate = true;
	}
	
	if (doDeallocate)
		DeallocateVisualData(visualPluginData);
	
	if (doAllocate)
		status = AllocateVisualData(visualPluginData, destRect);

	if (status != noErr)
		destPort = nil;

	visualPluginData->destPort = destPort;
	if (destRect != nil)
		visualPluginData->destRect = *destRect;

	return status;
}


/*
	ResetRenderData
*/
static void ResetRenderData (VisualPluginData *visualPluginData)
{
	ClearMemory(&visualPluginData->renderData, sizeof(visualPluginData->renderData));
	
	ClearMemory(&visualPluginData->waveformData[0][0], sizeof(visualPluginData->waveformData));
	
	visualPluginData->level[0] = 0;
	visualPluginData->level[1] = 0;
}

/*
void _writetolog(char *strData)
{
	FILE *fp;

	fp = fopen("c:\\myplugin.txt","a+");  
	fprintf(fp, strData);
	fprintf(fp, "\n");
	fclose(fp);
}
*/

/*
void _addtostring(char *strDest, char *strSource, int cbSource)
{
	char *sz = NULL;

	sz = malloc(cbSource + 1);
	strncpy(sz, strSource + 1, cbSource);

//	strDest = realloc(strDest, strlen(strDest) + strlen(sz) + 1);
//	strcat(strDest, sz);
	free(sz);

}
*/

DWORD _DoSnarl(char *Title, char *Text)
{
	char *base = "notify?app-sig=application/x-vnd.fullphat-sitp&uid=_itp_now_playing";
	char *t1 = "&title=";
	char *t2 = "&text=";
	char *req = NULL;
	DWORD hr;
	DWORD cb;

	hr = snDoRequest("register?app-sig=application/x-vnd.fullphat-sitp&title=iTunes");
	if (hr < 0)
		return(hr);

	/* size the buffer */

	cb = strlen(base);

	if (Title)
		cb = cb + strlen(t1) + strlen(Title);

	if (Text)
		cb = cb + strlen(t2) + strlen(Text);

	cb = cb + 1;
	req = malloc(cb);
	if (!req) 
		return(-666);

	/* clear the buffer */

	ClearMemory(req, cb);

	/* build the request string */

	strcpy(req, base);

	if (Title) {
		strcat(req, t1);
		strcat(req, Title);
	}

	if (Text) {
		strcat(req, t2);
		strcat(req, Text);
	}

	hr = snDoRequest(req);
	free(req);

	return(hr);

}


void _DoTrackChanged(ITTrackInfoV1 *trackInfo)
{
	char *lpName = NULL;
	char *lpText = NULL;
	char *lpAlbum = NULL;
	char *lpArtist = NULL;
	int  cb = 0;

	lpName = realloc(lpName, 255);
	ClearMemory(lpName, 255);
	strncpy(lpName, trackInfo->name + 1, trackInfo->name[0]);

//	if (trackInfo->album[0])
		cb = cb + 256;		//trackInfo->album[0];

//	if (cb)
		cb = cb + strlen(" by ");

//	if (trackInfo->artist[0])
		cb = cb + 256;

	cb = cb + 1;

	lpText = malloc(cb);
	ClearMemory(lpText, cb);

//	if (trackInfo->album[0]) {
		lpAlbum = malloc(256);		
		strncpy(lpAlbum, trackInfo->album + 1, 255);
		strcat(lpText, lpAlbum);
		free(lpAlbum);
//	}

	strcat(lpText, " by ");


//	if (trackInfo->artist[0]) {
		lpArtist = malloc(256);
		strncpy(lpArtist, trackInfo->artist + 1, 255);
		strcat(lpText, lpArtist);
		free(lpArtist);
//	}
	
	_DoSnarl(lpName, lpText);

	free(lpName);
	free(lpText);

}


/*
	VisualPluginHandler
*/
static OSStatus VisualPluginHandler (OSType message, VisualPluginMessageInfo *messageInfo, void *refCon)
{
	OSStatus			status;
	VisualPluginData *	visualPluginData;


	visualPluginData = (VisualPluginData *)refCon;
	
	status = noErr;

/*
	fp = fopen("c:\\myplugin.txt","a+");  
	fprintf(fp,"message=: %x \n",message);
	fclose(fp);
*/

	switch (message)
	{
		/*
			Sent when the visual plugin is registered.  The plugin should do minimal
			memory allocations here.  The resource fork of the plugin is still available.
		*/		
		case kVisualPluginInitMessage:
		{
			visualPluginData = (VisualPluginData *)malloc(sizeof(VisualPluginData));
			if (visualPluginData == nil)
			{
				status = memFullErr;
				break;
			}

			visualPluginData->appCookie	= messageInfo->u.initMessage.appCookie;
			visualPluginData->appProc	= messageInfo->u.initMessage.appProc;

			//messageInfo->u.initMessage.unused = kPluginWantsToBeLeftOpen;

			/* Remember the file spec of our plugin file. We need this so we can open our resource fork during */
			/* the configuration message */

			messageInfo->u.initMessage.refCon	= (void *)visualPluginData;
			break;
		}
			
		/*
			Sent when the visual plugin is unloaded
		*/		
		case kVisualPluginCleanupMessage:
			if (visualPluginData != nil)
				free(visualPluginData);
			break;
			
		/*
			Sent when the visual plugin is enabled.  iTunes currently enables all
			loaded visual plugins.  The plugin should not do anything here.
		*/
		case kVisualPluginEnableMessage:
		case kVisualPluginDisableMessage:
			break;

		/*
			Sent if the plugin requests idle messages.  Do this by setting the kVisualWantsIdleMessages
			option in the PlayerRegisterVisualPluginMessage.options field.
		*/
		case kVisualPluginIdleMessage:
			if (visualPluginData->playing == false)
				RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, false);
			break;
		
		/*
			Sent when iTunes is going to show the visual plugin in a port.  At
			this point, the plugin should allocate any large buffers it needs.
		*/
		case kVisualPluginShowWindowMessage:
			visualPluginData->destOptions = messageInfo->u.showWindowMessage.options;

			status = ChangeVisualPort(	visualPluginData,
										#if TARGET_OS_WIN32
											messageInfo->u.showWindowMessage.window,
										#endif
										&messageInfo->u.showWindowMessage.drawRect);
			if (status == noErr)
				RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, true);
			break;
			
		/*
			Sent when iTunes is no longer displayed.
		*/
		case kVisualPluginHideWindowMessage:
			(void) ChangeVisualPort(visualPluginData, nil, nil);

			ClearMemory(&visualPluginData->trackInfo, sizeof(visualPluginData->trackInfo));
			ClearMemory(&visualPluginData->streamInfo, sizeof(visualPluginData->streamInfo));
			break;
		
		/*
			Sent when iTunes needs to change the port or rectangle of the currently
			displayed visual.
		*/
		case kVisualPluginSetWindowMessage:
			visualPluginData->destOptions = messageInfo->u.setWindowMessage.options;

			status = ChangeVisualPort(	visualPluginData,
										#if TARGET_OS_WIN32
											messageInfo->u.setWindowMessage.window,
										#endif
										&messageInfo->u.setWindowMessage.drawRect);

			if (status == noErr)
				RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, true);
			break;
		
		/*
			Sent for the visual plugin to render a frame.
		*/
		case kVisualPluginRenderMessage:
			visualPluginData->renderTimeStampID	= messageInfo->u.renderMessage.timeStampID;

			ProcessRenderData(visualPluginData, messageInfo->u.renderMessage.renderData);
				
			RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, false);
			break;
			
		/*
			Sent in response to an update event.  The visual plugin should update
			into its remembered port.  This will only be sent if the plugin has been
			previously given a ShowWindow message.
		*/	
		case kVisualPluginUpdateMessage:
			RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, true);
			break;
		
		/*
			Sent when the player starts.
		*/
		case kVisualPluginPlayMessage:
			if (messageInfo->u.playMessage.trackInfo != nil) {
				visualPluginData->trackInfo = *messageInfo->u.playMessage.trackInfo;
				_DoTrackChanged(messageInfo->u.playMessage.trackInfo);

			}
			else
				ClearMemory(&visualPluginData->trackInfo, sizeof(visualPluginData->trackInfo));

			if (messageInfo->u.playMessage.streamInfo != nil)
				visualPluginData->streamInfo = *messageInfo->u.playMessage.streamInfo;
			else
				ClearMemory(&visualPluginData->streamInfo, sizeof(visualPluginData->streamInfo));
		
			visualPluginData->playing = true;
			break;

		/*
			Sent when the player changes the current track information.  This
			is used when the information about a track changes, or when the CD
			moves onto the next track.  The visual plugin should update any displayed
			information about the currently playing song.
		*/
		case kVisualPluginChangeTrackMessage:
			if (messageInfo->u.changeTrackMessage.trackInfo != nil) {
				visualPluginData->trackInfo = *messageInfo->u.changeTrackMessage.trackInfo;

			}
			else
				ClearMemory(&visualPluginData->trackInfo, sizeof(visualPluginData->trackInfo));

			if (messageInfo->u.changeTrackMessage.streamInfo != nil)
				visualPluginData->streamInfo = *messageInfo->u.changeTrackMessage.streamInfo;
			else
				ClearMemory(&visualPluginData->streamInfo, sizeof(visualPluginData->streamInfo));
			break;

		/*
			Sent when the player stops.
		*/
		case kVisualPluginStopMessage:
			visualPluginData->playing = false;
			_DoSnarl("iTunes", "Playback stopped");

			ResetRenderData(visualPluginData);

			RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, true);
			break;
		
		/*
			Sent when the player changes the track position.
		*/
		case kVisualPluginSetPositionMessage:
			break;

		/*
			Sent when the player pauses.  iTunes does not currently use pause or unpause.
			A pause in iTunes is handled by stopping and remembering the position.
		*/
		case kVisualPluginPauseMessage:
			visualPluginData->playing = false;

			ResetRenderData(visualPluginData);

			RenderVisualPort(visualPluginData, visualPluginData->destPort, &visualPluginData->destRect, true);
			break;
			
		/*
			Sent when the player unpauses.  iTunes does not currently use pause or unpause.
			A pause in iTunes is handled by stopping and remembering the position.
		*/
		case kVisualPluginUnpauseMessage:
			visualPluginData->playing = true;
			break;
		
		/*
			Sent to the plugin in response to a MacOS event.  The plugin should return noErr
			for any event it handles completely, or an error (unimpErr) if iTunes should handle it.
		*/
		case kVisualPluginEventMessage:
			status = unimpErr;
			break;

		default:
			status = unimpErr;
			break;
	}

	return status;	
}


/*
	RegisterVisualPlugin
*/
static OSStatus RegisterVisualPlugin (PluginMessageInfo *messageInfo)
{
	OSStatus			status;
	PlayerMessageInfo	playerMessageInfo;
		
	ClearMemory(&playerMessageInfo.u.registerVisualPluginMessage,sizeof(playerMessageInfo.u.registerVisualPluginMessage));
	
	// copy in name length byte first
	playerMessageInfo.u.registerVisualPluginMessage.name[0] = lstrlen(kSampleVisualPluginName);
	// now copy in actual name
	memcpy(&playerMessageInfo.u.registerVisualPluginMessage.name[1], kSampleVisualPluginName, lstrlen(kSampleVisualPluginName));

	SetNumVersion(&playerMessageInfo.u.registerVisualPluginMessage.pluginVersion, kSampleVisualPluginMajorVersion, kSampleVisualPluginMinorVersion, kSampleVisualPluginReleaseStage, kSampleVisualPluginNonFinalRelease);

	playerMessageInfo.u.registerVisualPluginMessage.options					= kVisualWantsIdleMessages | kVisualWantsConfigure;
	playerMessageInfo.u.registerVisualPluginMessage.handler					= VisualPluginHandler;
	playerMessageInfo.u.registerVisualPluginMessage.registerRefCon			= 0;
	playerMessageInfo.u.registerVisualPluginMessage.creator					= kSampleVisualPluginCreator;
	
	playerMessageInfo.u.registerVisualPluginMessage.timeBetweenDataInMS		= 0xFFFFFFFF; // 16 milliseconds = 1 Tick, 0xFFFFFFFF = Often as possible.
	playerMessageInfo.u.registerVisualPluginMessage.numWaveformChannels		= 2;
	playerMessageInfo.u.registerVisualPluginMessage.numSpectrumChannels		= 2;
	
	playerMessageInfo.u.registerVisualPluginMessage.minWidth				= 64;
	playerMessageInfo.u.registerVisualPluginMessage.minHeight				= 64;
	playerMessageInfo.u.registerVisualPluginMessage.maxWidth				= 32767;
	playerMessageInfo.u.registerVisualPluginMessage.maxHeight				= 32767;
	playerMessageInfo.u.registerVisualPluginMessage.minFullScreenBitDepth	= 0;
	playerMessageInfo.u.registerVisualPluginMessage.maxFullScreenBitDepth	= 0;
	playerMessageInfo.u.registerVisualPluginMessage.windowAlignmentInBytes	= 0;
	
	status = PlayerRegisterVisualPlugin(messageInfo->u.initMessage.appCookie, messageInfo->u.initMessage.appProc,&playerMessageInfo);
		
	return status;
	
}





/*
	MAIN
*/
IMPEXP OSStatus MAIN (OSType message, PluginMessageInfo *messageInfo, void *refCon)
{
	OSStatus		status;
	DWORD			hr;
	
	(void) refCon;
/*
	fp = fopen("c:\\myplugin.txt","a+");  
	fprintf(fp,"main(): message=: %d \n",message);
	fclose(fp);
*/
	switch (message)
	{
		case kPluginInitMessage:
			status = RegisterVisualPlugin(messageInfo);
			break;
			
		case kPluginCleanupMessage:
			hr = snDoRequest("unregister?app-sig=application/x-vnd.fullphat-sitp");
			status = noErr;
			break;
			
		default:
			status = unimpErr;
			break;
	}
	
	return status;
}




static HWND GetSnarlWindow()
{
	return FindWindow(SnarlWindowClass, SnarlWindowTitle);;
}


static DWORD snDoRequest(const char *request)
{
	DWORD nResult = 0;
	COPYDATASTRUCT cds;

	HWND hWnd = GetSnarlWindow();

	if (!IsWindow(hWnd))
		return -999;

	// Create COPYDATASTRUCT
	cds.dwData = 0x534E4C03;           // "SNL",3
	cds.cbData = (DWORD)strlen(request);
	cds.lpData = (char *)request;

	// Send message
	if (SendMessageTimeout(hWnd, WM_COPYDATA, (WPARAM)GetCurrentProcessId(), (LPARAM)&cds, SMTO_ABORTIFHUNG | SMTO_NOTIMEOUTIFNOTHUNG, 1000, &nResult) == 0)
	{
		DWORD nError = GetLastError();
		if (nError == ERROR_TIMEOUT)
			nResult = -901;
		else
			nResult = -951;
	}

	return nResult;
}





