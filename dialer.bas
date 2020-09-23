Attribute VB_Name = "Module1"

      Option Explicit
      #If Win32 Then
        Private Declare Function sndPlaySound Lib "winmm" Alias _
           "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) _
           As Long
      #Else
        Private Declare Function sndPlaySound Lib "MMSYSTEM" ( _
                           lpszSoundName As Any, ByVal uFlags%) As Integer
      #End If

'*********************************************************************

      ' Flag values for wFlags parameter.

'*********************************************************************

      Public Const SND_SYNC = &H0        ' Play synchronously (default).
      'Public Const SND_ASYNC = &H1      ' Play asynchronously (see
                                         ' note below).
      Public Const SND_NODEFAULT = &H2   ' Do not use default sound.
      Public Const SND_MEMORY = &H4      ' lpszSoundName points to a
                                         ' memory file.
      Public Const SND_LOOP = &H8        ' Loop the sound until next
                                         ' sndPlaySound.
      Public Const SND_NOSTOP = &H10     ' Do not stop any currently
                                         ' playing sound.

'*********************************************************************

      ' Plays a wave file from a resource.

'*********************************************************************
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

      Public Sub PlayWaveRes(vntResourceID As Variant, Optional vntFlags)
      '-----------------------------------------------------------------
      ' WARNING:  If you want to play sound files asynchronously in
      '           Win32, then you MUST change bytSound() from a local
      '           variable to a module-level or static variable. Doing
      '           this prevents your array from being destroyed before
      '           sndPlaySound is complete. If you fail to do this, you
      '           will pass an invalid memory pointer, which will cause
      '           a GPF in the Multimedia Control Interface (MCI).
      '-----------------------------------------------------------------
      Dim bytSound() As Byte ' Always store binary data in byte arrays!

      bytSound = LoadResData(vntResourceID, "WAVE")

      If IsMissing(vntFlags) Then
         vntFlags = SND_NODEFAULT Or SND_SYNC Or SND_MEMORY
      End If

      If (vntFlags And SND_MEMORY) = 0 Then
         vntFlags = vntFlags Or SND_MEMORY
      End If

      sndPlaySound bytSound(0), vntFlags
      End Sub




