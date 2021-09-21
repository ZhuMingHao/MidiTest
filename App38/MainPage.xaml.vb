' The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

Imports System.Runtime.InteropServices
''' <summary>
''' An empty page that can be used on its own or navigated to within a Frame.
''' </summary>
Public NotInheritable Class MainPage
    Inherits Page

    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim midiOut0 As New MidiOut2(0)
        'someNote you can use any number between 35 And 81, velocity = 127(max volume)
        midiOut0.Send(153, 55, 120)
    End Sub
End Class


Public Class MidiOut2

    <DllImport("winmm.dll", EntryPoint:="midiOutShortMsg")>
    Private Shared Function midiOutShortMsg(ByVal hMidiOut As IntPtr, ByVal msg As Integer) As UInteger

    End Function

    <DllImport("winmm.dll", EntryPoint:="midiOutOpen")>
    Private Shared Function midiOutOpen(ByRef lphMidiOut As IntPtr, ByVal uDeviceID%, ByVal dwCallback As IntPtr, ByVal dwInstance As IntPtr, ByVal dwFlags As Integer) As UInteger
    End Function

    Private hMidiOut As IntPtr = IntPtr.Zero
    Private shortMsg As Integer

    Public Sub Send(status%, toneNum%, vol%)
        shortMsg% = (vol << 16) Or (toneNum << 8) Or status
        midiOutShortMsg(hMidiOut, shortMsg)
    End Sub


    Public Sub New(DevID%)
        midiOutOpen(hMidiOut, DevID, Nothing, IntPtr.Zero, 0)
    End Sub


    <DllImport("winmm.dll", EntryPoint:="midiOutClose")>
    Private Shared Function midiOutclose(h As IntPtr) As UInteger
    End Function

    Public Sub Close()
        midiOutclose(hMidiOut)
    End Sub

End Class