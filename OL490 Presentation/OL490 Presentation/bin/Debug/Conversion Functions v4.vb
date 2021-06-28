Module ConversionFunctionsv3

    Structure DblValues
        Dim DblValue1, DblValue2 As Double
        Dim BoolSaturated As Boolean
    End Structure
    Structure DblValuesLuminance
        Dim DblValue1, DblValue2, DblLuminance As Double
        Dim BoolSaturated As Boolean
    End Structure
    Structure IntValues
        Dim IntValue1, IntValue2, IntValue3 As Integer
        Dim BoolSaturated As Boolean
    End Structure
    Structure Cone_Datum
        Dim IntWavelength As Integer
        Dim SngLMS() As Single
    End Structure
    Structure Cone_Data
        Dim Cones() As Cone_Datum
    End Structure
    Structure Spectrum_Datum
        Dim SngWavelength, SngIntensity As Single
    End Structure
    Structure Spectrum_Data
        Dim Spectrum() As Spectrum_Datum
    End Structure

    Public SngCIE1931(2, 1) As Single
    Public SngLum(2) As Single

#Region "Forward going conversions (Gun -> CIE '31 -> CIE '51 -> M-B -> Cart. -> Polar)"
    'Convert Gun %s to CIE 1931
    Function GunToCIE1931(ByVal DblOldRedPercent As Double, ByVal DblOldGreenPercent As Double) As DblValues
        Dim DblStep(8) As Double
        DblStep(0) = (SngCIE1931(1, 0) - SngCIE1931(2, 0)) * (SngCIE1931(0, 1) - SngCIE1931(2, 1)) - (SngCIE1931(1, 1) - SngCIE1931(2, 1)) * (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(1) = (SngCIE1931(0, 1) - SngCIE1931(2, 1)) / DblStep(0)
        DblStep(2) = -(SngCIE1931(0, 0) - SngCIE1931(2, 0)) / DblStep(0)
        DblStep(3) = ((-SngCIE1931(2, 0) * (SngCIE1931(0, 1) - SngCIE1931(2, 1))) + (SngCIE1931(2, 1) * (SngCIE1931(0, 0) - SngCIE1931(2, 0)))) / DblStep(0)
        DblStep(4) = (1 - ((SngCIE1931(1, 0) - SngCIE1931(2, 0)) * DblStep(1))) / (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(5) = (-DblStep(2) * (SngCIE1931(1, 0) - SngCIE1931(2, 0))) / (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(6) = (-SngCIE1931(2, 0) - (DblStep(3) * (SngCIE1931(1, 0) - SngCIE1931(2, 0)))) / (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(7) = (-SngCIE1931(1, 1) * DblStep(1) * DblStep(6) / DblStep(4)) + (SngCIE1931(1, 1) * DblStep(3))
        DblStep(8) = DblOldGreenPercent - (SngCIE1931(1, 1) * DblStep(1) / (SngCIE1931(0, 1) * DblStep(4))) * (DblOldRedPercent - (SngCIE1931(0, 1) * DblStep(5))) - (SngCIE1931(1, 1) * DblStep(2))

        GunToCIE1931.DblValue2 = DblStep(7) / DblStep(8) 'New y Value in CIE 1931
        GunToCIE1931.DblValue1 = (DblOldRedPercent * GunToCIE1931.DblValue2 - SngCIE1931(0, 1) * DblStep(5) * GunToCIE1931.DblValue2 - SngCIE1931(0, 1) * DblStep(6)) / (SngCIE1931(0, 1) * DblStep(4)) 'New x Value in CIE 1931
    End Function

    'Convert CIE 1931 to CIE 1951
    Function CIE1931ToCIE1951(ByVal DblOldX As Double, ByVal DblOldY As Double) As DblValues
        CIE1931ToCIE1951.DblValue1 = (1.0271 * DblOldX - 0.00008 * DblOldY - 0.00009) / (0.03845 * DblOldX + 0.01496 * DblOldY + 1) 'New x Value in CIE 1951
        CIE1931ToCIE1951.DblValue2 = (0.00376 * DblOldX + 1.0072 * DblOldY + 0.00764) / (0.03845 * DblOldX + 0.01496 * DblOldY + 1) 'New y Value in CIE 1951
    End Function

    'Convert CIE 1951 to MacLeod-Boynton
    Function CIE1951ToMB(ByVal DblOldX As Double, ByVal DblOldY As Double) As DblValues
        Dim DblStep(2) As Double
        DblStep(0) = 0.15514 * DblOldX + 0.54312 * DblOldY - 0.03286 * (1 - DblOldX - DblOldY)
        DblStep(1) = -0.15514 * DblOldX + 0.45684 * DblOldY + 0.03286 * (1 - DblOldX - DblOldY)
        DblStep(2) = 0.01608 * (1 - DblOldX - DblOldY)

        CIE1951ToMB.DblValue1 = DblStep(0) / (DblStep(0) + DblStep(1)) 'New L/(L+M) Value in MacLeod-Boynton
        CIE1951ToMB.DblValue2 = DblStep(2) / (DblStep(0) + DblStep(1)) 'New S/(L+M) Value in MacLeod-Boynton
    End Function

    'Convert MacLeod-Boynton to Cartesian
    Function MBToCartesian(ByVal DblOldL As Double, ByVal DblOldS As Double) As DblValues
        MBToCartesian.DblValue1 = 1955 * (DblOldL - 0.6568) 'New x Value in Cartesian
        MBToCartesian.DblValue2 = 5533 * (DblOldS - 0.01825) 'New y Value in Cartesian
    End Function

    'Convert Cartesian to Polar
    Function CartesianToPolar(ByVal DblOldX As Double, ByVal DblOldY As Double) As DblValues
        If DblOldX = 0 Then
            CartesianToPolar.DblValue1 = Math.Sign(DblOldY) * Math.PI / 2 'Plus or Minus 90 degrees (along the y axis)
        ElseIf DblOldX > 0 Then
            CartesianToPolar.DblValue1 = Math.Atan(DblOldY / DblOldX)
        Else 'DblOldX < 0
            CartesianToPolar.DblValue1 = Math.Atan(DblOldY / DblOldX) + Math.PI 'Adds 180 degrees to keep the point on the left side
        End If

        CartesianToPolar.DblValue1 = CartesianToPolar.DblValue1 * 180 / Math.PI 'Convert from Radians to Degrees, New Angle Value in Polar
        CartesianToPolar.DblValue2 = Math.Sqrt(DblOldX ^ 2 + DblOldY ^ 2) 'New Contrast Value in Polar
    End Function
#End Region 'No Try needed, not called from outside

#Region "Backward going conversions (Polar -> Cart. -> M-B -> CIE '51 -> CIE '31 -> Gun)"
    'Convert Polar to Cartesian
    Function PolarToCartesian(ByVal DblOldAngle As Double, ByVal DblOldContrast As Double) As DblValues
        PolarToCartesian.DblValue1 = DblOldContrast * Math.Cos(DblOldAngle * Math.PI / 180) 'New x Value in Cartesian
        PolarToCartesian.DblValue2 = DblOldContrast * Math.Sin(DblOldAngle * Math.PI / 180) 'New y Value in Cartesian
    End Function

    'Convert Cartesian to MacLeod-Boynton
    Function CartesianToMB(ByVal DblOldX As Double, ByVal DblOldY As Double) As DblValues
        CartesianToMB.DblValue1 = (DblOldX / 1955) + 0.6568 'New L/(L+M) Value in MacLeod-Boynton
        CartesianToMB.DblValue2 = (DblOldY / 5533) + 0.01825 'New S/(L+M) Value in MacLeod-Boynton
    End Function

    'Convert MacLeod-Boynton to CIE 1951
    Function MBToCIE1951(ByVal DblOldL As Double, ByVal DblOldS As Double) As DblValues
        Dim DblStep(2) As Double
        DblStep(0) = 2.94481 * DblOldL - 3.50097 * (1 - DblOldL) + 13.17218 * DblOldS
        DblStep(1) = 2.18895 * (1 - DblOldL) + 0.33959 * DblStep(0) - 4.47319 * DblOldS
        DblStep(2) = DblOldS / 0.01608

        MBToCIE1951.DblValue1 = DblStep(0) / (DblStep(0) + DblStep(1) + DblStep(2)) 'New x Value in CIE 1951
        MBToCIE1951.DblValue2 = DblStep(1) / (DblStep(0) + DblStep(1) + DblStep(2)) 'New y Value in CIE 1951
    End Function

    'Convert CIE 1951 to CIE 1931
    Function CIE1951ToCIE1931(ByVal DblOldX As Double, ByVal DblOldY As Double) As DblValues
        Dim DblStep(3) As Double
        DblStep(0) = (0.00376 - 0.03845 * DblOldY) / (1.0271 - 0.03845 * DblOldX)
        DblStep(1) = (0.00009 + DblOldX) * DblStep(0) + 0.00764 - DblOldY
        DblStep(2) = (0.03845 * DblOldY - 0.00376) / (1.0271 - 0.03845 * DblOldX)
        DblStep(3) = (0.00008 + 0.01496 * DblOldX) * DblStep(2) + 0.01496 * DblOldY - 1.0072

        CIE1951ToCIE1931.DblValue2 = DblStep(1) / DblStep(3) 'New y Value in CIE 1931
        CIE1951ToCIE1931.DblValue1 = (0.00008 * DblOldY + 0.00009 + (0.01496 * DblOldX * CIE1951ToCIE1931.DblValue2) + DblOldX) / (1.0271 - 0.03845 * DblOldX) 'New x Value in CIE 1931
    End Function

    'Convert CIE 1931 to Gun %s
    Function CIE1931ToGun(ByVal DblOldX As Double, ByVal DblOldY As Double) As DblValues
        Dim DblStep(5) As Double
        DblStep(0) = (DblOldX - SngCIE1931(2, 0)) * (SngCIE1931(0, 1) - SngCIE1931(2, 1)) - (DblOldY - SngCIE1931(2, 1)) * (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(1) = (SngCIE1931(1, 0) - SngCIE1931(2, 0)) * (SngCIE1931(0, 1) - SngCIE1931(2, 1)) - (SngCIE1931(1, 1) - SngCIE1931(2, 1)) * (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(2) = DblStep(0) / DblStep(1)
        DblStep(3) = ((DblOldX - SngCIE1931(2, 0)) - DblStep(2) * (SngCIE1931(1, 0) - SngCIE1931(2, 0))) / (SngCIE1931(0, 0) - SngCIE1931(2, 0))
        DblStep(4) = 1 - DblStep(2) - DblStep(3)
        DblStep(5) = DblStep(3) * SngCIE1931(0, 1) + DblStep(2) * SngCIE1931(1, 1) + DblStep(4) * SngCIE1931(2, 1)

        CIE1931ToGun.DblValue1 = DblStep(3) * SngCIE1931(0, 1) / DblStep(5) 'New Red Value in Gun %
        CIE1931ToGun.DblValue2 = DblStep(2) * SngCIE1931(1, 1) / DblStep(5) 'New Green Value in Gun %
    End Function
#End Region 'No Try needed, not called from outside

    Public Function ConvertColorSpace(ByVal IntFromSpace As Integer, ByVal IntToSpace As Integer, ByVal DblVal1 As Double, ByVal DblVal2 As Double) As DblValues
        Try
            With ConvertColorSpace
                .DblValue1 = DblVal1
                .DblValue2 = DblVal2
                'Forward going sequence
                If IntFromSpace = 0 And IntToSpace > 0 Then 'If starting at Gun %, convert first to CIE 1931
                    ConvertColorSpace = GunToCIE1931(.DblValue1, .DblValue2)
                End If
                If .DblValue1 <= 0 Or .DblValue2 <= 0 Or .DblValue1 >= 1 Or .DblValue2 >= 1 Then .BoolSaturated = True
                If IntFromSpace < 2 And IntToSpace >= 2 Then 'If starting at Gun % or CIE 1931, convert to CIE 1951
                    ConvertColorSpace = CIE1931ToCIE1951(.DblValue1, .DblValue2)
                End If
                If .DblValue1 <= 0 Or .DblValue2 <= 0 Or .DblValue1 >= 1 Or .DblValue2 >= 1 Then .BoolSaturated = True
                If IntFromSpace < 3 And IntToSpace >= 3 Then 'If starting at Gun %, CIE 1931, or CIE 1951, convert to MacLeod-Boynton
                    ConvertColorSpace = CIE1951ToMB(.DblValue1, .DblValue2)
                End If
                If .DblValue1 <= 0 Or .DblValue2 <= 0 Or .DblValue1 >= 1 Or .DblValue2 >= 1 Then .BoolSaturated = True
                If IntFromSpace < 4 And IntToSpace >= 4 Then 'If starting at Gun %, CIE 1931, CIE 1951, or MacLeod-Boynton, convert to Cartesian
                    ConvertColorSpace = MBToCartesian(.DblValue1, .DblValue2)
                End If
                If IntFromSpace < 5 And IntToSpace >= 5 Then 'If starting at anything prior to Polar, convert to Polar
                    ConvertColorSpace = CartesianToPolar(.DblValue1, .DblValue2)
                End If
                'Backward going sequence
                If IntFromSpace = 5 And IntToSpace < 5 Then 'If starting at Polar, convert to Cartesian
                    ConvertColorSpace = PolarToCartesian(.DblValue1, .DblValue2)
                End If
                If IntFromSpace >= 4 And IntToSpace < 4 Then 'If starting at Polar or Cartesian, convert to MacLeod-Boynton
                    ConvertColorSpace = CartesianToMB(.DblValue1, .DblValue2)
                End If
                If .DblValue1 < 0 Or .DblValue2 < 0 Or .DblValue1 > 1 Or .DblValue2 > 1 Then .BoolSaturated = True
                If IntFromSpace >= 3 And IntToSpace < 3 Then 'If starting at Polar, Cartesian, or MacLeod-Boynton, convert to CIE 1951
                    ConvertColorSpace = MBToCIE1951(.DblValue1, .DblValue2)
                End If
                If .DblValue1 <= 0 Or .DblValue2 <= 0 Or .DblValue1 >= 1 Or .DblValue2 >= 1 Then .BoolSaturated = True
                If IntFromSpace >= 2 And IntToSpace < 2 Then 'If starting at Polar, Cartesian, MB, or CIE1951, convert to 1931
                    ConvertColorSpace = CIE1951ToCIE1931(.DblValue1, .DblValue2)
                End If
                If .DblValue1 <= 0 Or .DblValue2 <= 0 Or .DblValue1 >= 1 Or .DblValue2 >= 1 Then .BoolSaturated = True
                If IntFromSpace > 0 And IntToSpace = 0 Then 'If starting at anything beoynd Gun %, convert to Gun %
                    ConvertColorSpace = CIE1931ToGun(.DblValue1, .DblValue2)
                    If 1 - .DblValue1 - .DblValue2 <= 0 Then .BoolSaturated = True
                End If
                If .DblValue1 <= 0 Or .DblValue2 <= 0 Or .DblValue1 >= 1 Or .DblValue2 >= 1 Then .BoolSaturated = True
            End With
        Catch ex As Exception
            MsgBox("Error converting color" & vbCrLf & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function RGBToGun(ByVal IntRed As Integer, ByVal IntGreen As Integer, ByVal IntBlue As Integer) As DblValuesLuminance
        Try
            Dim DblRGB(2), DblLuminance(2) As Double
            'Convert from 0 to 255 scale to 0 to 1 scale
            DblRGB(0) = IntRed / 255
            DblRGB(1) = IntGreen / 255
            DblRGB(2) = IntBlue / 255
            'Convert from RGB value to luminance value
            DblLuminance(0) = DblRGB(0) * SngLum(0)
            DblLuminance(1) = DblRGB(1) * SngLum(1)
            DblLuminance(2) = DblRGB(2) * SngLum(2)
            'Total to find luminance
            RGBToGun.DblLuminance = DblLuminance(0) + DblLuminance(1) + DblLuminance(2)
            'Convert from luminance value to gun %
            RGBToGun.DblValue1 = DblLuminance(0) / RGBToGun.DblLuminance
            RGBToGun.DblValue2 = DblLuminance(1) / RGBToGun.DblLuminance
            If RGBToGun.DblValue1 <= 0 Or RGBToGun.DblValue2 <= 0 Or RGBToGun.DblValue1 >= 1 Or RGBToGun.DblValue2 >= 1 Then RGBToGun.BoolSaturated = True
        Catch ex As Exception
            MsgBox("Error converting from RGB to Gun %" & vbCrLf & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function GunToRGB(ByVal DblOldRed As Double, ByVal DblOldGreen As Double, ByVal DblLuminance As Double) As IntValues
        Try
            Dim DblLum(2) As Double
            Dim IntIndex As Integer
            GunToRGB.BoolSaturated = False
            DblLum(0) = DblOldRed * DblLuminance 'Red luminance
            DblLum(1) = DblOldGreen * DblLuminance 'Green luminance
            DblLum(2) = (1 - DblOldRed - DblOldGreen) * DblLuminance 'Blue luminance
            For IntIndex = 0 To 2
                If DblLum(IntIndex) < 0 Then DblLum(IntIndex) = 0
            Next
            'Convert from Luminance value to 8-bit RGB
            If DblLum(0) < SngLum(0) Then
                GunToRGB.IntValue1 = 255 * (DblLum(0) / SngLum(0))
            Else
                GunToRGB.IntValue1 = 255 : GunToRGB.BoolSaturated = True
            End If
            If DblLum(1) < SngLum(1) Then
                GunToRGB.IntValue2 = 255 * (DblLum(1) / SngLum(1))
            Else
                GunToRGB.IntValue2 = 255 : GunToRGB.BoolSaturated = True
            End If
            If DblLum(2) < SngLum(2) Then
                GunToRGB.IntValue3 = 255 * (DblLum(2) / SngLum(2))
            Else
                GunToRGB.IntValue3 = 255 : GunToRGB.BoolSaturated = True
            End If
            'Prevent complete black from somehow turning to white (not sure how that was happening)
            If DblLuminance = 0 Then GunToRGB.IntValue1 = 0 : GunToRGB.IntValue2 = 0 : GunToRGB.IntValue3 = 0
        Catch ex As Exception
            MsgBox("Error converting from Gun % to RGB" & vbCrLf & vbCrLf & ex.Message)
        End Try
    End Function

    Public Function Open_Cone_Data(ByVal FileName As String) As Cone_Data
        ReDim Open_Cone_Data.Cones(0) 'Avoid NULL warning
        Dim IntDatum As Integer
        FileOpen(1, FileName, OpenMode.Input)
        While EOF(1) = False
            ReDim Preserve Open_Cone_Data.Cones(IntDatum)
            With Open_Cone_Data.Cones(IntDatum)
                ReDim .SngLMS(2)
                Input(1, .IntWavelength)
                Input(1, .SngLMS(0)) : Input(1, .SngLMS(1)) : Input(1, .SngLMS(2))
            End With
            IntDatum += 1
        End While
        FileClose(1)
    End Function

    Public Function Spectrum_to_MB(ByVal Spectrum As Spectrum_Data, ByVal ConeData As Cone_Data) As DblValues
        Dim SngWaveBounds(1) As Single : SngWaveBounds(0) = 380 : SngWaveBounds(1) = 825
        Dim SngLMS(2), SngInfluence(1) As Single
        Dim IntDatum, IntWav As Integer
        For IntDatum = 0 To Spectrum.Spectrum.Length - 1
            With Spectrum.Spectrum(IntDatum)
                Select Case .SngWavelength
                    Case Is <= SngWaveBounds(0) 'Ultraviolet
                        SngLMS(0) += ConeData.Cones(0).SngLMS(0)
                        SngLMS(1) += ConeData.Cones(0).SngLMS(1)
                        SngLMS(2) += ConeData.Cones(0).SngLMS(2)
                    Case Is >= SngWaveBounds(1) 'Infrared
                        SngLMS(0) += ConeData.Cones(ConeData.Cones.Length - 1).SngLMS(0)
                        SngLMS(1) += ConeData.Cones(ConeData.Cones.Length - 1).SngLMS(1)
                        SngLMS(2) += ConeData.Cones(ConeData.Cones.Length - 1).SngLMS(2)
                    Case Else 'Normal
                        IntWav = 0
                        Do
                            IntWav += 1
                        Loop Until .SngWavelength <= ConeData.Cones(IntWav).IntWavelength
                        SngInfluence(0) = 1 - Math.Abs((.SngWavelength - ConeData.Cones(IntWav - 1).IntWavelength) / (ConeData.Cones(IntWav).IntWavelength - ConeData.Cones(IntWav - 1).IntWavelength))
                        SngInfluence(1) = 1 - Math.Abs((.SngWavelength - ConeData.Cones(IntWav).IntWavelength) / (ConeData.Cones(IntWav).IntWavelength - ConeData.Cones(IntWav - 1).IntWavelength))
                        SngLMS(0) += .SngIntensity * ((SngInfluence(0) * ConeData.Cones(IntWav - 1).SngLMS(0)) + (SngInfluence(1) * ConeData.Cones(IntWav).SngLMS(0)))
                        SngLMS(1) += .SngIntensity * ((SngInfluence(0) * ConeData.Cones(IntWav - 1).SngLMS(1)) + (SngInfluence(1) * ConeData.Cones(IntWav).SngLMS(1)))
                        SngLMS(2) += .SngIntensity * ((SngInfluence(0) * ConeData.Cones(IntWav - 1).SngLMS(2)) + (SngInfluence(1) * ConeData.Cones(IntWav).SngLMS(2)))
                End Select
            End With
        Next
        Spectrum_to_MB.DblValue1 = SngLMS(0) / (SngLMS(0) + SngLMS(1))
        Spectrum_to_MB.DblValue2 = SngLMS(2) / (SngLMS(0) + SngLMS(1))
    End Function

End Module