Imports System.Drawing
Imports System.Numerics
Imports System.Threading.Tasks
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Runtime.CompilerServices

Public Module BitmapExtensions


    <Flags> _
    Public Enum ShakeBlurFlag
        ''' <summary>
        ''' 上下抖动
        ''' </summary>
        ''' <remarks></remarks>
        UpDown = 1
        ''' <summary>
        ''' 左右抖动
        ''' </summary>
        ''' <remarks></remarks>
        LeftRight = 2
        ''' <summary>
        ''' 四向抖动
        ''' </summary>
        ''' <remarks></remarks>
        All = 3
    End Enum

    Public Enum MotionBlurFlag
        Up = 1
        Down = 2
        Left = 3
        Right = 4
    End Enum

    Public Enum MirrorFlag
        Vertical = 1
        Horizontal = 2
    End Enum



    ''' <summary>
    ''' 二值化
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="color1"></param>
    ''' <param name="color2"></param>
    ''' <returns></returns>
    ''' <remarks>内部调用Gray</remarks>
    <Extension> _
    Public Function Binarization(ByVal asBitmap As Bitmap, Optional ByVal color1 As Color = Nothing, Optional ByVal color2 As Color = Nothing) As Bitmap
        If color1 = Nothing Then
            color1 = Color.Black
        End If
        If color2 = Nothing Then
            color2 = Color.White
        End If
        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim mGrayMap(255) As Integer
        Dim avgGray As Byte
        Dim sumGray, sum As UInt32
        bp = asBitmap.PsGray(mGrayMap)
        avgGray = 0 : sumGray = 0 : sum = 0
        For i = 0 To mGrayMap.Length - 1
            sumGray += mGrayMap(i) * i
            sum += mGrayMap(i)
        Next
        avgGray = sumGray / sum
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpBuffer.Length - 1 Step 4
            bpBuffer(i) = IIf(bpBuffer(i) >= avgGray, color2.B, color1.B)
            bpBuffer(i + 1) = IIf(bpBuffer(i + 1) >= avgGray, color2.G, color1.G)
            bpBuffer(i + 2) = IIf(bpBuffer(i + 2) >= avgGray, color2.R, color1.R)
        Next
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 剪裁
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="clipRect"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function Clip(ByVal asBitmap As Bitmap, ByVal clipRect As Rectangle) As Bitmap
        Dim bpRect As New Rectangle(0, 0, asBitmap.Width, asBitmap.Height)
        If Not bpRect.IntersectsWith(clipRect) Then
            Return asBitmap
        End If
        Dim bp As Bitmap
        Dim g As Graphics
        clipRect.Intersect(bpRect)
        bp = New Bitmap(clipRect.Width, clipRect.Height)
        g = Graphics.FromImage(bp)
        g.DrawImage(asBitmap, New Rectangle(0, 0, bp.Width, bp.Height), clipRect, GraphicsUnit.Pixel)
        g.Dispose()
        Return bp
    End Function

    ''' <summary>
    ''' 反色
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function ColorReverse(ByVal asBitmap As Bitmap) As Bitmap
        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpBuffer.Length - 1 Step 4
            bpBuffer(i) = 255 - bpBuffer(i)
            bpBuffer(i + 1) = 255 - bpBuffer(i + 1)
            bpBuffer(i + 2) = 255 - bpBuffer(i + 2)
        Next
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 对比度保留灰化
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function GrayEx(ByVal asBitmap As Bitmap) As Bitmap
        '-----缩放64*64-----
        Dim scale As Single = 64 / Math.Sqrt(asBitmap.Width * asBitmap.Height)
        Dim bp As New Bitmap(CInt(asBitmap.Width * scale), CInt(asBitmap.Height * scale))
        Dim g As Graphics = Graphics.FromImage(bp)
        g.InterpolationMode = InterpolationMode.NearestNeighbor
        g.DrawImage(asBitmap, New Rectangle(0, 0, bp.Width, bp.Height))
        g.Dispose()

        '-----灰度权重-----
        Dim sigma = 0.05!
        Dim sigma_pow As Single = sigma ^ 2
        Dim W As New List(Of Vector3)
        For i = 0 To 10
            For j = 0 To 10 - i
                Dim k = 10 - i - j
                W.Add(New Vector3(i / 10.0!, j / 10.0!, k / 10.0!))
            Next
        Next

        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim stride As Integer
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)
        stride = Math.Abs(bpData.Stride)
        ReDim bpBuffer(stride * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        bp.UnlockBits(bpData)

        '-----打乱像素-----
        Dim ran As Random
        Dim temp(bpBuffer.Length / 4 - 1) As Integer
        For i = 0 To temp.Length - 1
            temp(i) = i
        Next
        Dim shuffleBuffer(bpBuffer.Length - 1) As Byte
        For i = 0 To temp.Length - 1
            ran = New Random
            Dim pos = ran.Next(0, temp.Length - i)
            shuffleBuffer(i * 4) = bpBuffer(temp(pos) * 4)
            shuffleBuffer(i * 4 + 1) = bpBuffer(temp(pos) * 4 + 1)
            shuffleBuffer(i * 4 + 2) = bpBuffer(temp(pos) * 4 + 2)
            temp(pos) = temp(temp.Length - i - 1)
        Next

        '-----全局对比度-----
        Dim Delta As New List(Of Single)
        Dim P As New List(Of Vector3)
        For i = 0 To bpBuffer.Length - 1 Step 4
            Dim bB = CInt(bpBuffer(i)) - shuffleBuffer(i)
            Dim bG = CInt(bpBuffer(i + 1)) - shuffleBuffer(i + 1)
            Dim bR = CInt(bpBuffer(i + 2)) - shuffleBuffer(i + 2)
            Dim gB = bB / 255.0!
            Dim gG = bG / 255.0!
            Dim gR = bR / 255.0!
            Dim d = Math.Sqrt((bB ^ 2 + bG ^ 2 + bR ^ 2) >> 1) / 255.0!
            If d < sigma Then
                Continue For
            End If
            P.Add(New Vector3(gR, gG, gB))
            Delta.Add(d)
        Next

        '-----X轴对比度-----
        For j = 0 To bp.Height - 1
            For i = 0 To bp.Width - 2
                Dim bB = CInt(bpBuffer(j * stride + i * 4)) - bpBuffer(j * stride + (i + 1) * 4)
                Dim bG = CInt(bpBuffer(j * stride + i * 4 + 1)) - bpBuffer(j * stride + (i + 1) * 4 + 1)
                Dim bR = CInt(bpBuffer(j * stride + i * 4 + 2)) - bpBuffer(j * stride + (i + 1) * 4 + 2)
                Dim xB = bB / 255.0!
                Dim xG = bG / 255.0!
                Dim xR = bR / 255.0!
                Dim d = Math.Sqrt((bB ^ 2 + bG ^ 2 + bR ^ 2) >> 1) / 255.0!
                If d < sigma Then
                    Continue For
                End If
                P.Add(New Vector3(xR, xG, xB))
                Delta.Add(d)
            Next
        Next

        '-----Y轴对比度-----
        For i = 0 To bp.Width - 1
            For j = 0 To bp.Height - 2
                Dim bB = CInt(bpBuffer(j * stride + i * 4)) - bpBuffer((j + 1) * stride + i * 4)
                Dim bG = CInt(bpBuffer(j * stride + i * 4 + 1)) - bpBuffer((j + 1) * stride + i * 4 + 1)
                Dim bR = CInt(bpBuffer(j * stride + i * 4 + 2)) - bpBuffer((j + 1) * stride + i * 4 + 2)
                Dim yB = bB / 255.0!
                Dim yG = bG / 255.0!
                Dim yR = bR / 255.0!
                Dim d = Math.Sqrt((bB ^ 2 + bG ^ 2 + bR ^ 2) >> 1) / 255.0!
                If d < sigma Then
                    Continue For
                End If
                P.Add(New Vector3(yR, yG, yB))
                Delta.Add(d)
            Next
        Next

        Dim L(P.Count - 1, W.Count - 1) As Single
        Dim M1(P.Count - 1, W.Count - 1), M2(P.Count - 1, W.Count - 1) As Single
        For i = 0 To P.Count - 1
            For j = 0 To W.Count - 1
                L(i, j) = P(i).X * W(j).X + P(i).Y * W(j).Y + P(i).Z * W(j).Z
                M1(i, j) = L(i, j) + Delta(i)
                M2(i, j) = L(i, j) - Delta(i)
            Next
        Next

        Dim U(P.Count - 1, W.Count - 1) As Single
        For i = 0 To P.Count - 1
            For j = 0 To W.Count - 1
                U(i, j) = Math.Log(Math.E ^ (-(M1(i, j) ^ 2 / sigma_pow)) + Math.E ^ (-(M2(i, j) ^ 2 / sigma_pow)), 2)
            Next
        Next

        Dim E(W.Count - 1) As Single
        For j = 0 To W.Count - 1
            For i = 0 To P.Count - 1
                E(j) += U(i, j)
            Next
            E(j) /= P.Count
        Next

        Dim maxE As Single = E(0)
        Dim index As Integer = 0
        For i = 1 To E.Length - 1
            If maxE < E(i) Then
                maxE = E(i)
                index = i
            End If
        Next

        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpBuffer.Length - 1 Step 4
            Dim gray = W(index).X * bpBuffer(i + 2) + W(index).Y * bpBuffer(i + 1) + W(index).Z * bpBuffer(i)
            bpBuffer(i) = gray
            bpBuffer(i + 1) = gray
            bpBuffer(i + 2) = gray
        Next
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 灰化
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function Gray(ByVal asBitmap As Bitmap, Optional ByVal mGrayMap() As Integer = Nothing) As Bitmap
        Dim bp As Bitmap
        Dim bpBuffer() As Byte
        Dim bpData As BitmapData
        Dim R, G, B As Byte
        Dim bGray As Byte
        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpBuffer.Length - 1 Step 4
            B = bpBuffer(i)
            G = bpBuffer(i + 1)
            R = bpBuffer(i + 2)
            bGray = (R * 19595 + G * 38469 + B * 7472) >> 16
            If mGrayMap IsNot Nothing Then
                mGrayMap(bGray) += 1
            End If
            bpBuffer(i) = bGray : bpBuffer(i + 1) = bGray : bpBuffer(i + 2) = bGray
        Next
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 高斯模糊
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="radius">模糊半径</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function GaussianBlur(ByVal asBitmap As Bitmap, ByVal radius As Byte, Optional ByVal blurRect As Rectangle = Nothing) As Bitmap
        If radius < 2 Then
            Return asBitmap
        End If
        Dim bpRect As New Rectangle(0, 0, asBitmap.Width, asBitmap.Height)
        If blurRect = Nothing Then
            blurRect = bpRect
        End If
        If Not bpRect.IntersectsWith(blurRect) Then
            Return asBitmap
        End If
        blurRect.Intersect(bpRect)

        Dim startX, startY, endX, endY As UInt32
        startX = blurRect.X * 4 : endX = blurRect.Right * 4
        startY = blurRect.Y : endY = blurRect.Bottom

        Dim bp As Bitmap
        Dim bpBuffer() As Byte
        Dim tempBuffer() As Single
        Dim bpData As BitmapData
        Dim stride As Integer

        Dim gaussvalues(radius) As Single
        CalGauss(gaussvalues, radius)

        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        stride = Math.Abs(bpData.Stride)
        ReDim bpBuffer(stride * bpData.Height - 1)
        ReDim tempBuffer(bpBuffer.Length - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)

        Parallel.For(startY, endY, Sub(j As Integer)
                                       For i = startX To endX - 1 Step 4
                                           Dim sumGauss, gauss As Single
                                           Dim r, g, b, sumR, sumG, sumB As Single
                                           sumGauss = 0 : gauss = 0
                                           r = 0 : g = 0 : b = 0 : sumR = 0 : sumG = 0 : sumB = 0
                                           Dim rowoffset As UInt32 = stride * j
                                           For k = -radius To radius
                                               Dim xpos = i + 4 * k
                                               If xpos >= startX And xpos < endX Then
                                                   b = bpBuffer(rowoffset + xpos)
                                                   g = bpBuffer(rowoffset + xpos + 1)
                                                   r = bpBuffer(rowoffset + xpos + 2)
                                                   gauss = gaussvalues(Math.Abs(k))
                                                   sumB += b * gauss
                                                   sumG += g * gauss
                                                   sumR += r * gauss
                                                   sumGauss += gauss
                                               End If
                                           Next
                                           b = sumB / sumGauss
                                           g = sumG / sumGauss
                                           r = sumR / sumGauss
                                           tempBuffer(rowoffset + i) = b
                                           tempBuffer(rowoffset + i + 1) = g
                                           tempBuffer(rowoffset + i + 2) = r
                                       Next
                                   End Sub)

        Parallel.For(startX, endX, Sub(i As Integer)
                                       If i Mod 4 = 0 Then
                                           For j = startY To endY - 1
                                               Dim sumGauss, gauss As Single
                                               Dim r, g, b, sumR, sumG, sumB As Single
                                               sumGauss = 0 : gauss = 0
                                               r = 0 : g = 0 : b = 0 : sumR = 0 : sumG = 0 : sumB = 0
                                               Dim rowoffset As UInt32 = j * stride
                                               For k = -radius To radius
                                                   Dim ypos = j + k
                                                   If ypos >= startY And ypos < endY Then
                                                       Dim row_offset As UInt32 = ypos * stride
                                                       b = tempBuffer(row_offset + i)
                                                       g = tempBuffer(row_offset + i + 1)
                                                       r = tempBuffer(row_offset + i + 2)
                                                       gauss = gaussvalues(Math.Abs(k))
                                                       sumB += b * gauss
                                                       sumG += g * gauss
                                                       sumR += r * gauss
                                                       sumGauss += gauss
                                                   End If
                                               Next
                                               b = sumB / sumGauss
                                               g = sumG / sumGauss
                                               r = sumR / sumGauss
                                               bpBuffer(rowoffset + i) = Math.Min(b, 255)
                                               bpBuffer(rowoffset + i + 1) = Math.Min(g, 255)
                                               bpBuffer(rowoffset + i + 2) = Math.Min(r, 255)
                                           Next
                                       End If
                                   End Sub)

        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    Private Sub CalGauss(ByVal values() As Single, ByVal radius As Byte)
        Dim sigma As Single = radius / 3.0!
        Dim param1 As Single = 1 / (Math.Sqrt(2 * Math.PI) * sigma)
        Dim param2 As Single = -1 / (2 * sigma ^ 2)
        For i = 0 To values.Length - 1
            values(i) = param1 * Math.E ^ (i ^ 2 * param2)
        Next
    End Sub

    ''' <summary>
    ''' 镜像
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="flag"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function Mirror(ByVal asBitmap As Bitmap, ByVal flag As MirrorFlag) As Bitmap
        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Integer
        Dim tempBuffer() As Integer
        Dim stride As Integer

        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        stride = Math.Abs(bpData.Stride)
        ReDim bpBuffer(bpData.Width * bpData.Height - 1)
        ReDim tempBuffer(bpBuffer.Length - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpData.Width - 1
            For j = 0 To bpData.Height - 1
                Select Case flag
                    Case MirrorFlag.Vertical
                        tempBuffer(i + bpData.Width * j) = bpBuffer(i + bpData.Width * (bpData.Height - j - 1))
                    Case MirrorFlag.Horizontal
                        tempBuffer(i + bpData.Width * j) = bpBuffer(bpData.Width - i - 1 + bpData.Width * j)
                    Case Else
                        Return asBitmap
                End Select
            Next
        Next
        bpBuffer = tempBuffer
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 马赛克
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="mosaicRect">应用马赛克效果的矩形区域</param>
    ''' <param name="blocksize">区块大小(4～255)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function Mosaic(ByVal asBitmap As Bitmap, ByVal mosaicRect As Rectangle, ByVal blocksize As Byte) As Bitmap
        If blocksize < 4 Then
            Return asBitmap
        End If
        Dim bpRect As New Rectangle(0, 0, asBitmap.Width, asBitmap.Height)
        If Not bpRect.IntersectsWith(mosaicRect) Then
            Return asBitmap
        End If

        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim stride As Integer
        Dim bpBuffer() As Byte
        Dim r, g, b As Byte
        mosaicRect.Intersect(bpRect)
        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        stride = Math.Abs(bpData.Stride)
        ReDim bpBuffer(stride * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)

        For i = mosaicRect.X To mosaicRect.Right - 1 Step blocksize
            For j = mosaicRect.Y To mosaicRect.Bottom - 1 Step blocksize
                Dim offset = 4 * i + j * stride
                b = bpBuffer(offset)
                g = bpBuffer(offset + 1)
                r = bpBuffer(offset + 2)
                For m = 0 To blocksize - 1
                    Dim xpos = i + m
                    If xpos < bp.Width Then
                        For n = 0 To blocksize - 1
                            Dim ypos = j + n
                            If ypos < bp.Height Then
                                bpBuffer(4 * xpos + stride * ypos) = b
                                bpBuffer(4 * xpos + stride * ypos + 1) = g
                                bpBuffer(4 * xpos + stride * ypos + 2) = r
                            End If
                        Next
                    End If
                Next
            Next
        Next
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 运动模糊
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="offset">运动距离</param>
    ''' <param name="flag">运动方向</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function MotionBlur(ByVal asBitmap As Bitmap, ByVal offset As Byte, ByVal flag As MotionBlurFlag) As Bitmap
        If offset < 2 Then
            Return asBitmap
        End If

        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim tempBuffer() As Byte
        Dim stride As Integer
        Dim sumR, sumG, sumB, sum As Integer

        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        stride = Math.Abs(bpData.Stride)
        ReDim bpBuffer(stride * bpData.Height - 1)
        ReDim tempBuffer(bpBuffer.Length - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        tempBuffer = bpBuffer.Clone
        For i = 0 To bpData.Width - 1
            For j = 0 To bpData.Height - 1
                sumR = 0 : sumG = 0 : sumB = 0 : sum = 0
                For k = 0 To offset - 1
                    Select Case flag
                        Case MotionBlurFlag.Down
                            If j - k >= 0 Then
                                sumB += bpBuffer(4 * i + stride * (j - k))
                                sumG += bpBuffer(4 * i + stride * (j - k) + 1)
                                sumR += bpBuffer(4 * i + stride * (j - k) + 2)
                                sum += 1
                            End If
                        Case MotionBlurFlag.Up
                            If j + k < bpData.Height Then
                                sumB += bpBuffer(4 * i + stride * (j + k))
                                sumG += bpBuffer(4 * i + stride * (j + k) + 1)
                                sumR += bpBuffer(4 * i + stride * (j + k) + 2)
                                sum += 1
                            End If
                        Case MotionBlurFlag.Right
                            If i - k >= 0 Then
                                sumB += bpBuffer(4 * (i - k) + stride * j)
                                sumG += bpBuffer(4 * (i - k) + stride * j + 1)
                                sumR += bpBuffer(4 * (i - k) + stride * j + 2)
                                sum += 1
                            End If
                        Case MotionBlurFlag.Left
                            If i + k < bpData.Width Then
                                sumB += bpBuffer(4 * (i + k) + stride * j)
                                sumG += bpBuffer(4 * (i + k) + stride * j + 1)
                                sumR += bpBuffer(4 * (i + k) + stride * j + 2)
                                sum += 1
                            End If
                        Case Else
                            Return asBitmap
                    End Select
                Next
                tempBuffer(4 * i + stride * j) = sumB / sum
                tempBuffer(4 * i + stride * j + 1) = sumG / sum
                tempBuffer(4 * i + stride * j + 2) = sumR / sum
            Next
        Next

        bpBuffer = tempBuffer
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 灰化
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="mGreyMap"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function PsGray(ByVal asBitmap As Bitmap, Optional ByVal mGreyMap() As Integer = Nothing) As Bitmap
        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim r, g, b, max, min, mid As Byte
        Dim gray As Byte
        Dim rateMax, rateMaxMid As Single
        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpBuffer.Length - 1 Step 4
            b = bpBuffer(i) : g = bpBuffer(i + 1) : r = bpBuffer(i + 2)
            max = Math.Max(Math.Max(r, g), b)
            min = Math.Min(Math.Min(r, g), b)
            mid = CInt(r) + g + b - max - min
            Select Case max
                Case r, g
                    rateMax = 0.4!
                Case b
                    rateMax = 0.2!
            End Select
            Select Case min
                Case r, b
                    rateMaxMid = 0.6!
                Case g
                    rateMaxMid = 0.8!
            End Select
            gray = (max - mid) * rateMax + (mid - min) * rateMaxMid + min
            If mGreyMap IsNot Nothing Then
                mGreyMap(gray) += 1
            End If
            bpBuffer(i) = gray : bpBuffer(i + 1) = gray : bpBuffer(i + 2) = gray
        Next
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 旋转
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="angle">旋转角度(以度为单位)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function Rotate(ByVal asBitmap As Bitmap, ByVal angle As Integer) As Bitmap
        angle = angle Mod 360
        If angle = 0 Then
            Return asBitmap
        End If
        Dim rad As Single = angle / 180.0! * Math.PI
        Dim rotateRect As New Rectangle(0, 0, CInt(asBitmap.Width * Math.Abs(Math.Cos(rad)) + asBitmap.Height * Math.Abs(Math.Sin(rad))), CInt(asBitmap.Height * Math.Abs(Math.Cos(rad)) + asBitmap.Width * Math.Abs(Math.Sin(rad))))
        Dim bp As New Bitmap(rotateRect.Width, rotateRect.Height)
        Dim g As Graphics = Graphics.FromImage(bp)
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.TranslateTransform(rotateRect.Width / 2, rotateRect.Height / 2)
        g.RotateTransform(angle)
        g.TranslateTransform(-rotateRect.Width / 2, -rotateRect.Height / 2)
        g.DrawImage(asBitmap, New Rectangle((rotateRect.Width - asBitmap.Width) / 2, (rotateRect.Height - asBitmap.Height) / 2, asBitmap.Width, asBitmap.Height))
        g.Dispose()
        Return bp
    End Function

    ''' <summary>
    ''' HSL
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="nHue">色调(-180～180)</param>
    ''' <param name="nSaturation">饱和度(-100～100)</param>
    ''' <param name="nLightness">明度(-100～100)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function SetHSL(ByVal asBitmap As Bitmap, ByVal nHue As Integer, ByVal nSaturation As Integer, ByVal nLightness As Integer) As Bitmap
        If nHue = 0 And nSaturation = 0 And nLightness = 0 Then
            Return asBitmap
        End If
        nHue = Math.Min(Math.Max(nHue, -180), 180)
        nSaturation = Math.Min(Math.Max(nSaturation, -100), 100)
        nLightness = Math.Min(Math.Max(nLightness, -100), 100)
        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim hslBuffer() As Single
        Dim H, S, L, offsetH, offsetS, offsetL As Single
        offsetH = nHue / 180.0! * Math.PI
        offsetS = nSaturation / 100.0!
        offsetL = nLightness / 100.0!
        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        hslBuffer = RGB2HSL(bpBuffer)
        For i = 0 To hslBuffer.Length - 1 Step 4
            H = hslBuffer(i) + offsetH
            If H < 0 Then
                H += 2 * Math.PI
            ElseIf H >= 2 * Math.PI Then
                H -= 2 * Math.PI
            End If
            If offsetS <= 0 Then
                S = (1 + offsetS) * hslBuffer(i + 1)
            Else
                S = (1 - offsetS) * hslBuffer(i + 1) + offsetS
            End If
            If offsetL <= 0 Then
                L = (1 + offsetL) * hslBuffer(i + 2)
            Else
                L = (1 - offsetL) * hslBuffer(i + 2) + offsetL
            End If
            hslBuffer(i) = H
            hslBuffer(i + 1) = S
            hslBuffer(i + 2) = L
        Next
        bpBuffer = HSL2RGB(hslBuffer)
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' RGB2HSL
    ''' </summary>
    ''' <param name="bpBuffer">图像的ARGB信息</param>
    ''' <returns>返回图像的HSL信息(H:0～2PI S:0～1 L:0～1)</returns>
    ''' <remarks>传入的字节数组必须为BGRA形式,返回的HSL数组格式为HSLA</remarks>
    Public Function RGB2HSL(ByVal bpBuffer() As Byte) As Single()
        Dim sR, sG, sB, H, S, L As Single
        Dim delta As Single
        Dim max, min As Single
        Dim hslBuffer(bpBuffer.Length - 1) As Single
        For i = 0 To bpBuffer.Length - 1 Step 4
            sB = bpBuffer(i) / 255.0!
            sG = bpBuffer(i + 1) / 255.0!
            sR = bpBuffer(i + 2) / 255.0!
            max = Math.Max(Math.Max(sR, sG), sB)
            min = Math.Min(Math.Min(sR, sG), sB)
            delta = max - min
            If delta = 0 Then
                H = 0 : S = 0 : L = max
            Else
                Select Case max
                    Case sR
                        H = Math.PI / 3 * (((sG - sB) / delta) Mod 6)
                    Case sG
                        H = Math.PI / 3 * ((sB - sR) / delta + 2)
                    Case sB
                        H = Math.PI / 3 * ((sR - sG) / delta + 4)
                End Select
                L = (max + min) / 2
                S = delta / (1 - Math.Abs(2 * L - 1))
            End If
            hslBuffer(i) = H
            hslBuffer(i + 1) = S
            hslBuffer(i + 2) = L
            hslBuffer(i + 3) = bpBuffer(i + 3)
        Next
        Return hslBuffer
    End Function

    ''' <summary>
    ''' HSL2RGB
    ''' </summary>
    ''' <param name="hslBuffer"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HSL2RGB(ByVal hslBuffer() As Single) As Byte()
        Dim H, S, L As Single
        Dim c, x, m, nr, ng, nb As Single
        Dim bpBuffer(hslBuffer.Length - 1) As Byte
        For i = 0 To hslBuffer.Length - 1 Step 4
            H = hslBuffer(i) : S = hslBuffer(i + 1) : L = hslBuffer(i + 2)
            c = (1 - Math.Abs(2 * L - 1)) * S
            x = c * (1 - Math.Abs((H * 3 / Math.PI) Mod 2 - 1))
            m = L - c / 2
            Select Case H
                Case Is < Math.PI / 3
                    nr = c : ng = x : nb = 0
                Case Is < 2 * Math.PI / 3
                    nr = x : ng = c : nb = 0
                Case Is < Math.PI
                    nr = 0 : ng = c : nb = x
                Case Is < 4 * Math.PI / 3
                    nr = 0 : ng = x : nb = c
                Case Is < 5 * Math.PI / 3
                    nr = x : ng = 0 : nb = c
                Case Else
                    nr = c : ng = 0 : nb = x
            End Select
            bpBuffer(i) = (nb + m) * 255
            bpBuffer(i + 1) = (ng + m) * 255
            bpBuffer(i + 2) = (nr + m) * 255
            bpBuffer(i + 3) = hslBuffer(i + 3)
        Next
        Return bpBuffer
    End Function

    ''' <summary>
    ''' RGB2LAB
    ''' </summary>
    ''' <param name="bpBuffer"></param>
    ''' <returns></returns>
    ''' <remarks>D65-2</remarks>
    Public Function RGB2LAB(ByVal bpBuffer() As Byte) As Single()
        Dim sR, sG, sB As Single
        Dim X, Y, Z, L, A, B, fX, fY, fZ As Single
        Dim nX = 95.047!
        Dim nY = 100.0!
        Dim nZ = 108.883!
        Dim param1 = 0.008856!
        Dim param2 = 7.787!
        Dim param3 = 0.333333!
        Dim param4 = 0.1379!
        Dim labBuffer(bpBuffer.Length - 1) As Single
        For i = 0 To bpBuffer.Length - 1 Step 4
            '-----RGB2XYZ-----
            sB = bpBuffer(i) / 255.0! : sG = bpBuffer(i + 1) / 255.0! : sR = bpBuffer(i + 2) / 255.0!
            If sB > 0.04045! Then
                sB = ((sB + 0.055!) / 1.055!) ^ 2.4!
            Else
                sB = sB / 12.92!
            End If
            If sG > 0.04045! Then
                sG = ((sG + 0.055!) / 1.055!) ^ 2.4
            Else
                sG = sG / 12.92!
            End If
            If sR > 0.04045! Then
                sR = ((sR + 0.055!) / 1.055!) ^ 2.4
            Else
                sR = sR / 12.92!
            End If
            sB *= 100 : sG *= 100 : sR *= 100
            X = 0.4124! * sR + 0.3576! * sG + 0.1805! * sB
            Y = 0.2126! * sR + 0.7152! * sG + 0.0722! * sB
            Z = 0.0193! * sR + 0.1192! * sG + 0.9505! * sB
            '-----XYZ2LAB-----
            X /= nX : Y /= nY : Z /= nZ
            If X > param1 Then
                fX = X ^ param3
            Else
                fX = param2 * X + param4
            End If
            If Y > param1 Then
                fY = Y ^ param3
            Else
                fY = param2 * Y + param4
            End If
            If Z > param1 Then
                fZ = Z ^ param3
            Else
                fZ = param2 * Z + param4
            End If
            L = 116 * fY - 16
            A = 500 * (fX - fY)
            B = 200 * (fY - fZ)
            labBuffer(i) = L
            labBuffer(i + 1) = A
            labBuffer(i + 2) = B
            labBuffer(i + 3) = bpBuffer(i + 3)
        Next
        Return labBuffer
    End Function

    ''' <summary>
    ''' LAB2RGB
    ''' </summary>
    ''' <param name="labBuffer"></param>
    ''' <returns></returns>
    ''' <remarks>D65-2</remarks>
    Public Function LAB2RGB(ByVal labBuffer() As Single) As Byte()
        Dim L, A, B As Single
        Dim X, Y, Z, fX, fY, fZ As Single
        Dim sR, sG, sB As Single
        Dim nX = 95.047!
        Dim nY = 100.0!
        Dim nZ = 108.883!
        Dim param1 = 0.206897!
        Dim param2 = 0.137931!
        Dim param3 = 7.787!
        Dim param4 = 1 / 2.4!
        Dim bpBuffer(labBuffer.Length - 1) As Byte
        '------LAB2XYZ-----
        For i = 0 To labBuffer.Length - 1 Step 4
            L = labBuffer(i) : A = labBuffer(i + 1) : B = labBuffer(i + 2)
            fY = (L + 16) / 116
            fX = fY + A / 500
            fZ = fY - B / 200
            If fY > param1 Then
                fY = fY ^ 3
            Else
                fY = (fY - param2) / param3
            End If

            If fX > param1 Then
                fX = fX ^ 3
            Else
                fX = (fX - param2) / param3
            End If

            If fZ > param1 Then
                fZ = fZ ^ 3
            Else
                fZ = (fZ - param2) / param3
            End If
            X = fX * nX : Y = fY * nY : Z = fZ * nZ
            '-----XYZ2RGB-----
            X /= 100 : Y /= 100 : Z /= 100
            sR = 3.2406! * X - 1.5372! * Y - 0.4986! * Z
            sG = 1.8758! * Y - 0.9689! * X + 0.0415! * Z
            sB = 0.0557! * X - 0.204! * Y + 1.057! * Z
            If sR > 0.0031308! Then
                sR = 1.055! * (sR ^ param4) - 0.055!
            Else
                sR = 12.92! * sR
            End If

            If sG > 0.0031308! Then
                sG = 1.055! * (sG ^ param4) - 0.055!
            Else
                sG = 12.92! * sG
            End If

            If sB > 0.0031308! Then
                sB = 1.055! * (sB ^ param4) - 0.055!
            Else
                sB = 12.92! * sB
            End If
            sB *= 255 : sG *= 255 : sR *= 255
            bpBuffer(i) = Math.Min(Math.Max(sB, 0), 255)
            bpBuffer(i + 1) = Math.Min(Math.Max(sG, 0), 255)
            bpBuffer(i + 2) = Math.Min(Math.Max(sR, 0), 255)
            bpBuffer(i + 3) = labBuffer(i + 3)
        Next
        Return bpBuffer
    End Function

    ''' <summary>
    ''' 等比例缩放
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="nScale">缩放比</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function Scale(ByVal asBitmap As Bitmap, ByVal nScale As Single) As Bitmap
        If nScale <= 0 Or nScale = 1 Then
            Return asBitmap
        End If
        Dim bp As New Bitmap(CInt(asBitmap.Width * nScale), CInt(asBitmap.Height * nScale))
        Dim g As Graphics = Graphics.FromImage(bp)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(asBitmap, New Rectangle(0, 0, bp.Width, bp.Height))
        g.Dispose()
        Return bp
    End Function

    ''' <summary>
    ''' 泛黄
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function SepiaTone(ByVal asBitmap As Bitmap) As Bitmap
        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim r, g, b As Single
        Dim blend As Single
        Dim ran As New Random

        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        ReDim bpBuffer(Math.Abs(bpData.Stride) * bpData.Height - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        For i = 0 To bpBuffer.Length - 1 Step 4
            b = (139 * bpBuffer(i + 2) + 273 * bpBuffer(i + 1) + 67 * bpBuffer(i)) >> 9
            g = (179 * bpBuffer(i + 2) + 351 * bpBuffer(i + 1) + 86 * bpBuffer(i)) >> 9
            r = (201 * bpBuffer(i + 2) + 394 * bpBuffer(i + 1) + 97 * bpBuffer(i)) >> 9
            b = Math.Min(b, 255)
            g = Math.Min(g, 255)
            r = Math.Min(r, 255)
            blend = ran.NextDouble * 0.5! + 0.5!
            b = b * blend + bpBuffer(i) * (1 - blend)
            g = g * blend + bpBuffer(i + 1) * (1 - blend)
            r = r * blend + bpBuffer(i + 2) * (1 - blend)
            bpBuffer(i) = b
            bpBuffer(i + 1) = g
            bpBuffer(i + 2) = r
        Next

        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

    ''' <summary>
    ''' 抖动模糊
    ''' </summary>
    ''' <param name="asBitmap"></param>
    ''' <param name="offset">抖动距离</param>
    ''' <param name="flag">抖动类型</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension> _
    Public Function ShakeBlur(ByVal asBitmap As Bitmap, ByVal offset As Byte, ByVal flag As ShakeBlurFlag) As Bitmap
        If offset = 0 Then
            Return asBitmap
        End If

        Dim bp As Bitmap
        Dim bpData As BitmapData
        Dim bpBuffer() As Byte
        Dim tempBuffer() As Byte
        Dim stride As Integer
        Dim sumR, sumG, sumB, sum As Integer

        bp = asBitmap.Clone
        bpData = bp.LockBits(New Rectangle(0, 0, bp.Width, bp.Height), ImageLockMode.ReadWrite, PixelFormat.Format32bppArgb)
        stride = Math.Abs(bpData.Stride)
        ReDim bpBuffer(stride * bpData.Height - 1), tempBuffer(bpBuffer.Length - 1)
        Marshal.Copy(bpData.Scan0, bpBuffer, 0, bpBuffer.Length)
        tempBuffer = bpBuffer.Clone
        For i = 0 To bpData.Width - 1
            For j = 0 To bpData.Height - 1
                sumB = 0 : sumG = 0 : sumR = 0 : sum = 0
                Dim up, down, left, right As Integer
                up = j - offset : down = j + offset : left = i - offset : right = i + offset

                If flag.HasFlag(ShakeBlurFlag.UpDown) Then
                    If up >= 0 Then
                        sumB += bpBuffer(4 * i + stride * up)
                        sumG += bpBuffer(4 * i + stride * up + 1)
                        sumR += bpBuffer(4 * i + stride * up + 2)
                        sum += 1
                    End If
                    If down < bpData.Height Then
                        sumB += bpBuffer(4 * i + stride * down)
                        sumG += bpBuffer(4 * i + stride * down + 1)
                        sumR += bpBuffer(4 * i + stride * down + 2)
                        sum += 1
                    End If
                End If

                If flag.HasFlag(ShakeBlurFlag.LeftRight) Then
                    If left >= 0 Then
                        sumB += bpBuffer(4 * left + stride * j)
                        sumG += bpBuffer(4 * left + stride * j + 1)
                        sumR += bpBuffer(4 * left + stride * j + 2)
                        sum += 1
                    End If
                    If right < bpData.Width Then
                        sumB += bpBuffer(4 * right + stride * j)
                        sumG += bpBuffer(4 * right + stride * j + 1)
                        sumR += bpBuffer(4 * right + stride * j + 2)
                        sum += 1
                    End If
                End If
                tempBuffer(4 * i + stride * j) = sumB / sum
                tempBuffer(4 * i + stride * j + 1) = sumG / sum
                tempBuffer(4 * i + stride * j + 2) = sumR / sum
            Next
        Next

        bpBuffer = tempBuffer
        Marshal.Copy(bpBuffer, 0, bpData.Scan0, bpBuffer.Length)
        bp.UnlockBits(bpData)
        Return bp
    End Function

End Module
