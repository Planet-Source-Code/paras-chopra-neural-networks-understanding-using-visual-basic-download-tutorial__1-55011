Attribute VB_Name = "modNN"
'Don't forget to write option base 1 into the code
' or else this net will not work

'Coded by Paras Chopra
'paraschopra@lycos.com
'http://naramcheez.netfirms.com

'Please don't forget to give comments, credits and most important your VOTE!

Option Base 1
Option Explicit

Const e = 2.7183 'Mathematical const, used in sigmod function

Private Type Dendrite ' Dendrite connects one neuron to another and allows signal to pass from it
Weight As Double 'Weight it has
End Type

Private Type Neuron 'The main thing
Dendrites() As Dendrite 'Array of Denrites
DendriteCount As Long 'Number of dendrites
Bias As Double 'The bias
Value As Double 'The value to be passed to next layer of neurons
Delta As Double 'The delta of neuron (used while learning)
End Type



Private Type Layer 'Layer contaning number of neurons
Neurons() As Neuron 'Neurons in the layer
NeuronCount As Long 'Number of neurons
End Type

Private Type NeuralNetwork
Layers() As Layer 'Layers in the network
LayerCount As Long 'Number of layers
LearningRate As Double 'The learning rateof the network
End Type

Dim Network As NeuralNetwork ' Our main network

Function CreateNet(LearningRate As Double, ArrayOfLayers As Variant) As Integer '0 = Unsuccesful and 1 = Successful
Dim i, j, k As Integer
Network.LayerCount = UBound(ArrayOfLayers) 'Init number of layers
If Network.LayerCount < 2 Then 'Input and output layers must be there
    CreateNet = 0 'Unsuccessful
    Exit Function
End If
Network.LearningRate = LearningRate 'The learning rate
ReDim Network.Layers(Network.LayerCount) As Layer 'Redim the layers variable
For i = 1 To UBound(ArrayOfLayers) ' Initialize all layers
DoEvents
    Network.Layers(i).NeuronCount = ArrayOfLayers(i)
    ReDim Network.Layers(i).Neurons(Network.Layers(i).NeuronCount) As Neuron
    For j = 1 To ArrayOfLayers(i) 'Initialize all neurons
    DoEvents
        If i = UBound(ArrayOfLayers) Then 'We will not init dendrites for it because output layers doesn't have any
            Network.Layers(i).Neurons(j).Bias = GetRand 'Set the bias to random value
            Network.Layers(i).Neurons(j).DendriteCount = ArrayOfLayers(i - 1)
            ReDim Network.Layers(i).Neurons(j).Dendrites(Network.Layers(i).Neurons(j).DendriteCount) As Dendrite 'Redim the dendrite var
            For k = 1 To ArrayOfLayers(i - 1)
                DoEvents
                Network.Layers(i).Neurons(j).Dendrites(k).Weight = GetRand 'Set the weight of each dendrite
            Next k
        ElseIf i = 1 Then 'Only init dendrites not bias
            DoEvents 'Do nothing coz it is input layer
        Else
            Network.Layers(i).Neurons(j).Bias = GetRand 'Set the bias to random value
            Network.Layers(i).Neurons(j).DendriteCount = ArrayOfLayers(i - 1)
            ReDim Network.Layers(i).Neurons(j).Dendrites(Network.Layers(i).Neurons(j).DendriteCount) As Dendrite 'Redim the dendrite var
            For k = 1 To ArrayOfLayers(i - 1)
                DoEvents
                Network.Layers(i).Neurons(j).Dendrites(k).Weight = GetRand 'Set the weight of each dendrite
            Next k
        End If
    Next j
Next i
CreateNet = 1
End Function


Function Run(ArrayOfInputs As Variant) As Variant 'It returns the output inf form of array
Dim i, j, k As Integer
If UBound(ArrayOfInputs) <> Network.Layers(1).NeuronCount Then
    Run = 0
    Exit Function
End If
For i = 1 To Network.LayerCount
DoEvents
    For j = 1 To Network.Layers(i).NeuronCount
    DoEvents
        If i = 1 Then
            Network.Layers(i).Neurons(j).Value = ArrayOfInputs(j) 'Set the value of input layer
        Else
            Network.Layers(i).Neurons(j).Value = 0 'First set the value to zero
            For k = 1 To Network.Layers(i - 1).NeuronCount
                DoEvents
                Network.Layers(i).Neurons(j).Value = Network.Layers(i).Neurons(j).Value + Network.Layers(i - 1).Neurons(k).Value * Network.Layers(i).Neurons(j).Dendrites(k).Weight 'Calculating the value
            Next k
        Network.Layers(i).Neurons(j).Value = Activation(Network.Layers(i).Neurons(j).Value + Network.Layers(i).Neurons(j).Bias) 'Calculating the real value of neuron
        'Network.Layers(i).Neurons(j).Value = tanh(Network.Layers(i).Neurons(j).Value + Network.Layers(i).Neurons(j).Bias) 'Calculating the real value of neuron
        End If
    Next j
Next i
ReDim OutputResult(Network.Layers(Network.LayerCount).NeuronCount) As Double
For i = 1 To (Network.Layers(Network.LayerCount).NeuronCount)
    DoEvents
    OutputResult(i) = (Network.Layers(Network.LayerCount).Neurons(i).Value) 'The array of output result
Next i
Run = OutputResult
End Function

Function SupervisedTrain(inputdata As Variant, outputdata As Variant) As Integer '0=unsuccessful and 1 = sucessful
Dim i, j, k As Integer
If UBound(inputdata) <> Network.Layers(1).NeuronCount Then 'Check if correct amount of input is given
    SupervisedTrain = 0
    Exit Function
End If
If UBound(outputdata) <> Network.Layers(Network.LayerCount).NeuronCount Then 'Check if correct amount of output is given
    SupervisedTrain = 0
    Exit Function
End If
Call Run(inputdata) 'Calculate values of all neurons and set the input
'Calculate delta's
For i = 1 To Network.Layers(Network.LayerCount).NeuronCount
DoEvents
    Network.Layers(Network.LayerCount).Neurons(i).Delta = Network.Layers(Network.LayerCount).Neurons(i).Value * (1 - Network.Layers(Network.LayerCount).Neurons(i).Value) * (outputdata(i) - Network.Layers(Network.LayerCount).Neurons(i).Value) 'Deltas of Output layer
    For j = Network.LayerCount - 1 To 2 Step -1
    DoEvents
        For k = 1 To Network.Layers(j).NeuronCount
        DoEvents
            Network.Layers(j).Neurons(k).Delta = Network.Layers(j).Neurons(k).Value * (1 - Network.Layers(j).Neurons(k).Value) * Network.Layers(j + 1).Neurons(i).Dendrites(k).Weight * Network.Layers(j + 1).Neurons(i).Delta 'Deltas of Hidden Layers
        Next k
    Next j
Next i


For i = Network.LayerCount To 2 Step -1
DoEvents
    For j = 1 To Network.Layers(i).NeuronCount
    DoEvents
        Network.Layers(i).Neurons(j).Bias = Network.Layers(i).Neurons(j).Bias + (Network.LearningRate * 1 * Network.Layers(i).Neurons(j).Delta)  'Calculate new bias
        For k = 1 To Network.Layers(i).Neurons(j).DendriteCount
        DoEvents
            
            Network.Layers(i).Neurons(j).Dendrites(k).Weight = Network.Layers(i).Neurons(j).Dendrites(k).Weight + (Network.LearningRate * Network.Layers(i - 1).Neurons(k).Value * Network.Layers(i).Neurons(j).Delta) 'Calculate new weights
        Next k
    Next j
Next i
SupervisedTrain = 1
End Function


'Function Sigmod(Value As Double, Threshold As Double)
'Sigmod = 1 / (1 + e ^ (-(Value - Threshold)))
'End Function

'Using tanh instead of sigmod

Function tanh(x As Double) As Double
tanh = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function


Private Function Activation(Value As Double)
'To crunch a number between 0 and 1
    Activation = (1 / (1 + Exp(Value * -1)))
End Function

Function GetRand() As Double 'Produces a number between -1 and 1
Randomize
GetRand = 2 - (1 + Rnd + Rnd)
'GetRand = Rnd
End Function

Sub EraseNetwork()
Erase Network.Layers
End Sub

Function SaveNet(FilePath As String) As Integer ' 1 = successful, 0 =unsucessful
Dim i, j, k As Integer
Open FilePath For Output As #1
Print #1, "START Learning Rate"
Print #1, Network.LearningRate
Print #1, "END Learning Rate"
Print #1, "START Layer Count"
Print #1, Network.LayerCount
Print #1, "END Layer Count"
Print #1, "START Input Layer Neuron Count"
Print #1, Network.Layers(1).NeuronCount
Print #1, "END Input Layer Neuron Count"
For i = 2 To Network.LayerCount
    Print #1, "START Next Layer"
    Print #1, "START Neuron Count"
    Print #1, Network.Layers(i).NeuronCount
    Print #1, "END Neuron Count"
    For j = 1 To Network.Layers(i).NeuronCount
        Print #1, "START Neuron"
        Print #1, "START Bias"
        Print #1, Network.Layers(i).Neurons(j).Bias
        Print #1, "END Bias"
        Print #1, "START Dendrites"
        For k = 1 To Network.Layers(i).Neurons(j).DendriteCount
            Print #1, Network.Layers(i).Neurons(j).Dendrites(k).Weight
        Next k
        Print #1, "END Dendrites"
        Print #1, "END Neuron"
    Next j
    Print #1, "END Layer"
Next i
Close #1
SaveNet = 1
End Function

Function LoadNet(FilePath As String) As Integer ' 1 = successful, 0 =unsucessful
Dim Data, DataMain As String
Dim LayerTrack, NeuronTrack As Long 'The variable which tracks the current layer and current neuron
Dim i As Long
If FileExists(FilePath) = 0 Then
    LoadNet = 0 'File doest not exists
    Exit Function
End If
Open FilePath For Input As #1
Do While Not EOF(1)
    DoEvents
    Line Input #1, Data
    Select Case Data
        Case "START Learning Rate":
            Line Input #1, DataMain
            Network.LearningRate = CDbl(DataMain)
        Case "START Layer Count":
            Line Input #1, DataMain
            Network.LayerCount = CLng(DataMain)
            ReDim Network.Layers(Network.LayerCount) As Layer
        Case "START Input Layer Neuron Count": 'Input layer
            LayerTrack = 1
            Line Input #1, DataMain
            Network.Layers(1).NeuronCount = CLng(DataMain)
            ReDim Network.Layers(1).Neurons(Network.Layers(1).NeuronCount) As Neuron
        Case "START Neuron Count":
            LayerTrack = LayerTrack + 1
            Line Input #1, DataMain
            Network.Layers(LayerTrack).NeuronCount = CLng(DataMain)
            ReDim Network.Layers(LayerTrack).Neurons(Network.Layers(LayerTrack).NeuronCount) As Neuron
        Case "START Bias":
            NeuronTrack = NeuronTrack + 1
            Line Input #1, DataMain
            Network.Layers(LayerTrack).Neurons(NeuronTrack).Bias = CDbl(DataMain)
            Network.Layers(LayerTrack).Neurons(NeuronTrack).DendriteCount = Network.Layers(LayerTrack - 1).NeuronCount
            ReDim Network.Layers(LayerTrack).Neurons(NeuronTrack).Dendrites(Network.Layers(LayerTrack).Neurons(NeuronTrack).DendriteCount) As Dendrite
        Case "START Dendrites":
            For i = 1 To Network.Layers(LayerTrack).Neurons(NeuronTrack).DendriteCount 'All the dendrites
                DoEvents
                Line Input #1, DataMain
                Network.Layers(LayerTrack).Neurons(NeuronTrack).Dendrites(i).Weight = CDbl(DataMain)
            Next i
        Case "END Layer":
            NeuronTrack = 0
        Case Else
            DoEvents
    End Select
Loop
Close #1
LayerTrack = 0
NeuronTrack = 0
LoadNet = 1
End Function

' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Private Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)

    Close intFileNum

    Err = 0
End Function


Function UnSupervisedTrain(inputdata As Variant, outputdata As Variant) As Integer '0=unsuccessful and 1 = sucessful

End Function
