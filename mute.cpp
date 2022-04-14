/*
 Mute
           Copyright (c) 2022, Alexander Steinhoefer

-----------------------------------------------------------------------------
Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright notice,
      this list of conditions and the following disclaimer.

    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.

    * Neither the name of the author nor the names of its contributors may
      be used to endorse or promote products derived from this software
      without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.
-----------------------------------------------------------------------------
*/

#ifndef _WIN32
#  error Only windows is supported
#endif

#include <cstdlib>
#include <cstdio>
#include <cstring>
#include <cstdarg>
#include <memory>
#include <string>

#define NO_GDI
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <comip.h>
#include <comdef.h>
#include <Mmdeviceapi.h>
#include <audiopolicy.h>
#include <endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>


/* =============================================================================
 *  Types
 */

struct Options {
   bool silent;
   bool unmute;
};

_COM_SMARTPTR_TYPEDEF(IPropertyStore, __uuidof(IPropertyStore));
_COM_SMARTPTR_TYPEDEF(IMMDevice, __uuidof(IMMDevice));
_COM_SMARTPTR_TYPEDEF(IMMDeviceCollection, __uuidof(IMMDeviceCollection));
_COM_SMARTPTR_TYPEDEF(IAudioSessionManager2, __uuidof(IAudioSessionManager2));
_COM_SMARTPTR_TYPEDEF(IAudioEndpointVolume, __uuidof(IAudioEndpointVolume));
_COM_SMARTPTR_TYPEDEF(IMMDeviceEnumerator, __uuidof(IMMDeviceEnumerator));
_COM_SMARTPTR_TYPEDEF(IAudioSessionControl, __uuidof(IAudioSessionControl));

/* =============================================================================
 *  Globals
 */

static const char* programName_ = nullptr;
static struct Options opts_ = { 0 };

/* =============================================================================
 *  Output
 */

static void PrintError(const wchar_t* fmt, ...)
{
   if (opts_.silent) {
      return;
   }
   va_list ap;
   va_start(ap, fmt);
   fputs("! ", stdout);
   vfwprintf(stderr, fmt, ap);
   fputs("\n", stderr);
   va_end(ap);
}

static void Print(const wchar_t* fmt, ...)
{
   if (opts_.silent) {
      return;
   }
   va_list ap;
   va_start(ap, fmt);
   vwprintf(fmt, ap);
   fputs("\n", stdout);
   va_end(ap);
}

/* =============================================================================
 *  Mute
 */

static void MuteEndpoint(
   IAudioEndpointVolumePtr ev,
   const std::wstring& deviceName)
{
   BOOL isMuted = FALSE;
   HRESULT hr = ev->GetMute(&isMuted);
   if (FAILED(hr)) {
      PrintError(
         L"Failed to get mute status for device \"%ls\"",
         deviceName.c_str());
      return;
   }
   if (opts_.unmute && !isMuted) {
      Print(L"> %ls is already unmuted.", deviceName.c_str());
      return;
   } else if (!opts_.unmute && isMuted) {
      Print(L"> %ls is already muted.", deviceName.c_str());
      return;
   }

   hr = ev->SetMute(!opts_.unmute, nullptr);
   if (FAILED(hr)) {
      PrintError(
         L"Failed to set mute status for device \"%ls\"",
         deviceName.c_str());
   } else {
      Print(
         L"> %ls is now %lsmuted",
         deviceName.c_str(),
         (opts_.unmute) ? L"un" : L"");
   }
}

static bool Mute()
{
   IMMDeviceEnumeratorPtr deviceEnumerator;

   if (FAILED(deviceEnumerator.CreateInstance(
         __uuidof(MMDeviceEnumerator),
         nullptr,
         CLSCTX_INPROC_SERVER))) {
      PrintError(L"Failed to create instance of MMDeviceEnumerator");
      return false;
   }

   IMMDeviceCollectionPtr audioEndpoints;
   HRESULT hr = deviceEnumerator->EnumAudioEndpoints(
      eRender,
      DEVICE_STATE_ACTIVE,
      &audioEndpoints);
   if (FAILED(hr)) {
      PrintError(L"Failed to enumerate all audio endpoints");
      return false;
   }

   UINT epCount;
   hr = audioEndpoints->GetCount(&epCount);
   if (FAILED(hr)) {
      PrintError(L"Failed to get endpoint count");
      return false;
   }

   for (UINT i = 0; i < epCount; ++i) {
      IMMDevicePtr device = nullptr;
      IPropertyStorePtr propStore;

      hr = audioEndpoints->Item(i, &device);
      if (FAILED(hr)) {
         PrintError(L"Failed to get audio endpoint #%d", i);
         continue;
      }

      hr = device->OpenPropertyStore(STGM_READ, &propStore);
      if (FAILED(hr)) {
         PrintError(L"Failed to open property store for audio endpoint #%d", i);
         continue;
      }

      PROPVARIANT value;
      PropVariantInit(&value);
      hr = propStore->GetValue(PKEY_Device_FriendlyName, &value);
      if (FAILED(hr)) {
         PrintError(L"Failed to get device name for audio endpoint #%d");
         continue;
      }
      std::wstring deviceName = value.pwszVal;
      PropVariantClear(&value);

      Print(L"Found audio endpoint \"%ls\"", deviceName.c_str());

      IAudioSessionManager2Ptr sessionManager2;
      if (FAILED(device->Activate(
            __uuidof(IAudioSessionManager2), CLSCTX_INPROC_SERVER, nullptr,
            reinterpret_cast<LPVOID*>(&sessionManager2)))) {
         PrintError(
            L"Failed to retrieve audio session manager for \"%ls\"",
            deviceName.c_str());
         continue;
      }

      IAudioSessionControlPtr sessionCtrl;
      hr = sessionManager2->GetAudioSessionControl(nullptr, 0, &sessionCtrl);
      if (FAILED(hr)) {
         PrintError(
            L"Failed to retrieve audio session manager for \"%ls\"",
            deviceName.c_str());
         continue;
      }

      IAudioEndpointVolumePtr endpointVolume;
      hr = device->Activate(
         __uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER,
         nullptr, reinterpret_cast<LPVOID*>(&endpointVolume));
      if (FAILED(hr)) {
         PrintError(
            L"Failed to active endpoint volume for device \"%ls\"",
            deviceName.c_str());
         continue;
      }

      MuteEndpoint(endpointVolume, deviceName);
      fputs("\n", stdout);
   }

   return true;
}

/* =============================================================================
 *  Main and Command Line
 */

static void PrintUsage()
{
   std::printf(
      "%s <options>\n"
      "Options:\n"
      "\t-help\tDisplay this screen and exits\n"
      "\t-silent\tDon't print any output\n"
      "\t-unmute\tinstead of muting, do the opposite\n",
      programName_);
}

static bool DisplayUsage(int argc, char** argv)
{
   if (argc == 2 && (!_strcmpi(argv[1], "-help") || !strcmp(argv[1], "-?"))) {
      return true;
   }
   return false;
}

static bool ParseCommandLine(int argc, char** argv)
{
   for (int i = 1; i < argc; ++i) {
      if (_strcmpi(argv[i], "-silent") == 0) {
         opts_.silent = 1;
      } else if (_strcmpi(argv[i], "-unmute") == 0) {
         opts_.unmute = 1;
      } else {
         return false;
      }
   }
   return true;
}


static bool Init(int argc, char** argv)
{
   UNREFERENCED_PARAMETER(argc);
   if ((programName_ = strrchr(argv[0], '\\')) != nullptr) {
      programName_ += 1;
   } else {
      programName_ = argv[0];
   }
   if (CoInitializeEx(0, COINIT_APARTMENTTHREADED) != S_OK) {
      PrintError(L"Failed to initialize COM library");
      return false;
   }
   return true;
}

static void Shutdown()
{
   CoUninitialize();
}

int main(int argc, char** argv)
{
   int rc = EXIT_FAILURE;
   if (Init(argc, argv)) {
      if (DisplayUsage(argc, argv) || !ParseCommandLine(argc, argv)) {
         PrintUsage();
      } else if (Mute()) {
         rc = EXIT_SUCCESS;
      }
      Shutdown();
   }
   return rc;
}
