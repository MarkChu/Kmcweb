<%
'-----------------------------------------------------
' 描述: Asp無件上傳進度條
' 作者: 寶玉(www.webuc.net)
' 連結: www.pspsoft.com, www.cnforums.net, blog.joycode.com, www.cnblogs.com, www.51js.com
' 版本: 1.0 Beta
' 版權: 本作品可免費使用，但是勿移除版權信息
' 推薦：asp.net上傳件(http://www.upload4asp.net/)
'-----------------------------------------------------

Dim DoteyUpload_SourceData

Class DoteyUpload
	
	Public Files
	Public Form
	Public MaxTotalBytes
	Public Version
	Public ProgressID
	Public ErrMsg
	
	Private BytesRead
	Private ChunkReadSize
	Private Info
	Private Progress

	Private UploadProgressInfo
	Private CrLf

	Private Sub Class_Initialize()
		Set Files = Server.CreateObject("Scripting.Dictionary")	' 上傳文件集合
		Set Form = Server.CreateObject("Scripting.Dictionary")	' 表單集合
		UploadProgressInfo = "DoteyUploadProgressInfo"  ' Application的Key
		MaxTotalBytes = 1 *1024 *1024 *1024 ' 預設最大1G
		ChunkReadSize = 64 * 1024	' 分塊大小64K
		CrLf = Chr(13) & Chr(10)	' 換行

		Set DoteyUpload_SourceData = Server.CreateObject("ADODB.Stream")
		DoteyUpload_SourceData.Type = 1 ' 二進制流
		DoteyUpload_SourceData.Open

		Version = "1.0 Beta"	' 版本
		ErrMsg = ""	' 錯誤訊息
		Set Progress = New ProgressInfo

	End Sub

	' 將文件依據其文件名統一保存在某路徑下
	Public Sub SaveTo(path)
		
		Upload()	' 上傳

		if right(path,1) <> "/" then path = path & "/" 

		' 遍歷所有已上傳文件
		For Each fileItem In Files.Items			
			fileItem.SaveAs path & fileItem.FileName
		Next

		' 保存束后更新進度信息
		Progress.ReadyState = "complete" '上傳束
		UpdateProgressInfo progressID

	End Sub

	' 分析上傳的數據，並保存到相關集合中
	Public Sub Upload ()

		Dim TotalBytes, Boundary
		TotalBytes = Request.TotalBytes	 ' 總大小
		If TotalBytes < 1 Then
			Raise("無數據傳入")
			Exit Sub
		End If
		If TotalBytes > MaxTotalBytes Then
			Raise("您目前上傳的大小為" & TotalBytes/1000 & " K，最大允許為" & MaxTotalBytes/1024 & "K")
			Exit Sub
		End If
		Boundary = GetBoundary()
		If IsNull(Boundary) Then 
			Raise("如果form中有包括multipart/form-data上傳是無效的")
			Exit Sub	 ''如果form中有包括multipart/form-data上傳是無效的
		End If
		Boundary = StringToBinary(Boundary)
		
		Progress.ReadyState = "loading" '開始上傳
		Progress.TotalBytes = TotalBytes
		UpdateProgressInfo progressID

		Dim DataPart, PartSize
		BytesRead = 0

		'循環分塊讀取
		Do While BytesRead < TotalBytes

			'分塊讀取
			PartSize = ChunkReadSize
			if PartSize + BytesRead > TotalBytes Then PartSize = TotalBytes - BytesRead
			DataPart = Request.BinaryRead(PartSize)
			BytesRead = BytesRead + PartSize

			DoteyUpload_SourceData.Write DataPart

			Progress.UploadedBytes = BytesRead
			Progress.LastActivity = Now()

			' 更新進度信息
			UpdateProgressInfo progressID

		Loop

		' 上傳束后更新進度信息
		Progress.ReadyState = "loaded" '上傳束
		UpdateProgressInfo progressID

		Dim Binary
		DoteyUpload_SourceData.Position = 0
		Binary = DoteyUpload_SourceData.Read

		Dim BoundaryStart, BoundaryEnd, PosEndOfHeader, IsBoundaryEnd
		Dim Header, bFieldContent
		Dim FieldName
		Dim File
		Dim TwoCharsAfterEndBoundary

		BoundaryStart = InStrB(Binary, Boundary)
		BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary, 0)

		Do While (BoundaryStart > 0 And BoundaryEnd > 0 And Not IsBoundaryEnd)
			' 獲取表單頭的結束位置
			PosEndOfHeader = InStrB(BoundaryStart + LenB(Boundary), Binary, StringToBinary(vbCrLf + vbCrLf))
						
			' 分離表單頭信息,類似於：
			' Content-Disposition: form-data; name="file1"; filename="G:\homepage.txt"
			' Content-Type: text/plain
			Header = BinaryToString(MidB(Binary, BoundaryStart + LenB(Boundary) + 2, PosEndOfHeader - BoundaryStart - LenB(Boundary) - 2))

			' 分離表單內容
			bFieldContent = MidB(Binary, (PosEndOfHeader + 4), BoundaryEnd - (PosEndOfHeader + 4) - 2)
			
			FieldName = GetFieldName(Header)
			' 如果是附件
			If InStr (Header,"filename=""") > 0 Then
				Set File = New FileInfo
				
				' 獲取文件相關信息
				Dim clientPath
				clientPath = GetFileName(Header)
				File.FileName = GetFileNameByPath(clientPath)
				File.FileExt = GetFileExt(clientPath)
				File.FilePath = clientPath
				File.FileType = GetFileType(Header)
				File.FileStart = PosEndOfHeader + 3
				File.FileSize = BoundaryEnd - (PosEndOfHeader + 4) - 2
				File.FormName = FieldName

				' 如果該文件不為空並不存在該表單項保存之
				If Not Files.Exists(FieldName) And File.FileSize > 0 Then
					Files.Add FieldName, File
				End If
			'表單數據				
			Else
				' 允許同名表單
				If Form.Exists(FieldName) Then
					Form(FieldName) = Form(FieldName) & "," & BinaryToString(bFieldContent)
				Else
					Form.Add FieldName, BinaryToString(bFieldContent)
				End If
			End If

			' 是否束位置
			TwoCharsAfterEndBoundary = BinaryToString(MidB(Binary, BoundaryEnd + LenB(Boundary), 2))
			IsBoundaryEnd = TwoCharsAfterEndBoundary = "--"

			If Not IsBoundaryEnd Then ' 如果不是尾, 繼續讀取下一塊
				BoundaryStart = BoundaryEnd
				BoundaryEnd = InStrB(BoundaryStart + LenB(Boundary), Binary, Boundary)
			End If
		Loop
		
		' 解析文件束後更新進度信息
		Progress.UploadedBytes = TotalBytes
		Progress.ReadyState = "interactive" '解析文件束
		UpdateProgressInfo progressID

	End Sub

	'異常訊息
	Private Sub Raise(Message)
		ErrMsg = ErrMsg & "[" & Now & "]" & Message & "<BR>"
		
		Progress.ErrorMessage = Message
		UpdateProgressInfo ProgressID
		
		'call Err.Raise(vbObjectError, "DoteyUpload", Message)

	End Sub

	' 取邊界值
	Private Function GetBoundary()
		Dim ContentType, ctArray, bArray
		ContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
		ctArray = Split(ContentType, ";")
		If Trim(ctArray(0)) = "multipart/form-data" Then
			bArray = Split(Trim(ctArray(1)), "=")
			GetBoundary = "--" & Trim(bArray(1))
		Else	'如果form中有包括multipart/form-data上傳是無效的
			GetBoundary = null
			Raise("如果form中有包括multipart/form-data上傳是無效的")
		End If
	End Function

	' 將二進制轉化成文本
	Private Function BinaryToString(xBinary)
		Dim Binary
		if vartype(xBinary) = 8 then Binary = MultiByteToBinary(xBinary) else Binary = xBinary
		
	  Dim RS, LBinary
	  Const adLongVarChar = 201
	  Set RS = CreateObject("ADODB.Recordset")
	  LBinary = LenB(Binary)
		
		if LBinary>0 then
			RS.Fields.Append "mBinary", adLongVarChar, LBinary
			RS.Open
			RS.AddNew
				RS("mBinary").AppendChunk Binary 
			RS.Update
			BinaryToString = RS("mBinary")
		Else
			BinaryToString = ""
		End If
	End Function


	Function MultiByteToBinary(MultiByte)
	  Dim RS, LMultiByte, Binary
	  Const adLongVarBinary = 205
	  Set RS = CreateObject("ADODB.Recordset")
	  LMultiByte = LenB(MultiByte)
		if LMultiByte>0 then
			RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
			RS.Open
			RS.AddNew
				RS("mBinary").AppendChunk MultiByte & ChrB(0)
			RS.Update
			Binary = RS("mBinary").GetChunk(LMultiByte)
		End If
	  MultiByteToBinary = Binary
	End Function


	' 字串符到二進制
	Function StringToBinary(String)
		Dim I, B
		For I=1 to len(String)
			B = B & ChrB(Asc(Mid(String,I,1)))
		Next
		StringToBinary = B
	End Function

	'返回表單名
	Private Function GetFieldName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "name=")
		EndPos = InStr(sPos + 6, infoStr, Chr(34) & ";")
		If EndPos = 0 Then
			EndPos = inStr(sPos + 6, infoStr, Chr(34))
		End If
		GetFieldName = Mid(infoStr, sPos + 6, endPos - _
			(sPos + 6))
	End Function

	'返回文件名
	Private Function GetFileName(infoStr)
		Dim sPos, EndPos
		sPos = InStr(infoStr, "filename=")
		EndPos = InStr(infoStr, Chr(34) & CrLf)
		GetFileName = Mid(infoStr, sPos + 10, EndPos - _
			(sPos + 10))
	End Function

	'返回文件的 MIME type
	Private Function GetFileType(infoStr)
		sPos = InStr(infoStr, "Content-Type: ")
		GetFileType = Mid(infoStr, sPos + 14)
	End Function

	'根據路徑獲取文件名
	Private Function GetFileNameByPath(FullPath)
		Dim pos
		pos = 0
		FullPath = Replace(FullPath, "/", "\")
		pos = InStrRev(FullPath, "\") + 1
		If (pos > 0) Then
			GetFileNameByPath = Mid(FullPath, pos)
		Else
			GetFileNameByPath = FullPath
		End If
	End Function

	'根據路徑獲取擴展名
	Private Function GetFileExt(FullPath)
		Dim pos
		pos = InStrRev(FullPath,".")
		if pos>0 then GetFileExt = Mid(FullPath, Pos)
	End Function

	' 更新進度信息
	' 進度信息保存在Application中的ADODB.Recordset對象中
	Private Sub UpdateProgressInfo(progressID)
		Const adTypeText = 2, adDate = 7, adUnsignedInt = 19, adVarChar = 200
		
		If (progressID <> "" And IsNumeric(progressID)) Then
			Application.Lock()
			if IsEmpty(Application(UploadProgressInfo)) Then
				Set Info = Server.CreateObject("ADODB.Recordset")
				Set Application(UploadProgressInfo) = Info
				Info.Fields.Append "ProgressID", adUnsignedInt
				Info.Fields.Append "StartTime", adDate
				Info.Fields.Append "LastActivity", adDate
				Info.Fields.Append "TotalBytes", adUnsignedInt
				Info.Fields.Append "UploadedBytes", adUnsignedInt
				Info.Fields.Append "ReadyState", adVarChar, 128
				Info.Fields.Append "ErrorMessage", adVarChar, 4000
				Info.Open 
		 		Info("ProgressID").Properties("Optimize") = true
				Info.AddNew 
			Else
				Set Info = Application(UploadProgressInfo)
				If Not Info.Eof Then
					Info.MoveFirst()
					Info.Find "ProgressID = " & progressID
				End If
				If (Info.EOF) Then
					Info.AddNew
				End If
			End If

			Info("ProgressID") = clng(progressID)
			Info("StartTime") = Progress.StartTime
			Info("LastActivity") = Now()
			Info("TotalBytes") = Progress.TotalBytes
			Info("UploadedBytes") = Progress.UploadedBytes
			Info("ReadyState") = Progress.ReadyState
			Info("ErrorMessage") = Progress.ErrorMessage
			Info.Update

			Application.UnLock
		End IF
	End Sub

	' 根據上傳ID獲取進度信息
	Public Function GetProgressInfo(progressID)

		Dim pi, Infos
		Set pi = New ProgressInfo
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Set Infos = Application(UploadProgressInfo)
			If Not Infos.Eof Then
				Infos.MoveFirst
				Infos.Find "ProgressID = " & progressID
				If Not Infos.EOF Then
					pi.StartTime = Infos("StartTime")
					pi.LastActivity = Infos("LastActivity")
					pi.TotalBytes = clng(Infos("TotalBytes"))
					pi.UploadedBytes = clng(Infos("UploadedBytes"))
					pi.ReadyState = Trim(Infos("ReadyState"))
					pi.ErrorMessage = Trim(Infos("ErrorMessage"))
					Set GetProgressInfo = pi
				End If
			End If
		End If
		Set GetProgressInfo = pi
	End Function

	' 移除指定的進度信息
	Private Sub RemoveProgressInfo(progressID)
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Application.Lock
			Set Info = Application(UploadProgressInfo)
			If Not Info.Eof Then
				Info.MoveFirst
				Info.Find "ProgressID = " & progressID
				If  Not Info.EOF Then
					Info.Delete
				End If
			End If

			' 如果沒有記錄了, 直接釋放, 避免'800a0bcd'錯誤
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

	' 移除指定的進度信息
	Private Sub RemoveOldProgressInfo(progressID)
		If Not IsEmpty(Application(UploadProgressInfo)) Then
			Dim L
			Application.Lock

			Set Info = Application(UploadProgressInfo)
			Info.MoveFirst

			Do
				L = Info("LastActivity").Value
				If IsEmpty(L) Then
					Info.Delete() 
				ElseIf DateDiff("d", Now(), L) > 30 Then
					Info.Delete()
				End If
				Info.MoveNext()
			Loop Until Info.EOF

			' 如果沒有記錄了, 直接釋放, 避免'800a0bcd'錯誤
			If Info.RecordCount = 0 Then
				Info.Close
				Application.Contents.Remove UploadProgressInfo
			End If
			Application.UnLock
		End If
	End Sub

End Class

'---------------------------------------------------
' 進度信息 類
'---------------------------------------------------
Class ProgressInfo
	
	Public UploadedBytes
	Public TotalBytes
	Public StartTime
	Public LastActivity
	Public ReadyState
	Public ErrorMessage

	Private Sub Class_Initialize()
		UploadedBytes = 0	' 已上傳大小
		TotalBytes = 0	' 總大小
		StartTime = Now()	' 開始時間
		LastActivity = Now()	 ' 最後更新時間
		ReadyState = "uninitialized"	' uninitialized,loading,loaded,interactive,complete
		ErrorMessage = ""
	End Sub

	' 總大小
	Public Property Get TotalSize
		TotalSize = FormatNumber(TotalBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' 已上傳大小
	Public Property Get SizeCompleted
		SizeCompleted = FormatNumber(UploadedBytes / 1024, 0, 0, 0, -1) & " K"
	End Property 

	' 已上傳秒數
	Public Property Get ElapsedSeconds
		ElapsedSeconds = DateDiff("s", StartTime, Now())
	End Property 

	' 已上傳時間
	Public Property Get ElapsedTime
		If ElapsedSeconds > 3600 then
			ElapsedTime = ElapsedSeconds \ 3600 & " 時 " & (ElapsedSeconds mod 3600) \ 60 & " 分 " & ElapsedSeconds mod 60 & " 秒"
		ElseIf ElapsedSeconds > 60 then
			ElapsedTime = ElapsedSeconds \ 60 & " 分 " & ElapsedSeconds mod 60 & " 秒"
		else
			ElapsedTime = ElapsedSeconds mod 60 & " 秒"
		End If
	End Property 

	' 傳輸速率
	Public Property Get TransferRate
		If ElapsedSeconds > 0 Then
			TransferRate = FormatNumber(UploadedBytes / 1024 / ElapsedSeconds, 2, 0, 0, -1) & " K/秒"
		Else
			TransferRate = "0 K/秒"
		End If
	End Property 

	' 完成百分比
	Public Property Get Percentage
		If TotalBytes > 0 Then
			Percentage = fix(UploadedBytes / TotalBytes * 100) & "%"
		Else
			Percentage = "0%"
		End If
	End Property 

	' 估計剩餘時間
	Public Property Get TimeLeft
		If UploadedBytes > 0 Then
			SecondsLeft = fix(ElapsedSeconds * (TotalBytes / UploadedBytes - 1))
			If SecondsLeft > 3600 then
				TimeLeft = SecondsLeft \ 3600 & " 時 " & (SecondsLeft mod 3600) \ 60 & " 分 " & SecondsLeft mod 60 & " 秒"
			ElseIf SecondsLeft > 60 then
				TimeLeft = SecondsLeft \ 60 & " 分 " & SecondsLeft mod 60 & " 秒"
			else
				TimeLeft = SecondsLeft mod 60 & " 秒"
			End If
		Else
			TimeLeft = "未知"
		End If
	End Property 

End Class

'---------------------------------------------------
' 文件信息 類
'---------------------------------------------------
Class FileInfo
	
	Dim FormName, FileName, FilePath, FileSize, FileType, FileStart, FileExt, NewFileName

	Private Sub Class_Initialize 
		FileName = ""		' 文件名
		FilePath = ""			' Client端路徑
		FileSize = 0			' 文件大小
		FileStart= 0			' 文件開始位置
		FormName = ""	'表單名
		FileType = ""		' 文件Content Type
		FileExt = ""			'文件擴展名
		NewFileName = ""	'上傳後文件名稱
	End Sub

	Public Function Save()
		SaveAs(FileName)
	End Function

	'保存文件
	Public Function SaveAs(fullpath)
		Dim dr
		SaveAs = false
		If trim(fullpath) = "" Or FileStart = 0 Or FileName = "" Or right(fullpath,1) = "/" Then Exit Function
		
		NewFileName = GetFileNameByPath(fullpath)

		Set dr = CreateObject("Adodb.Stream")
		dr.Mode = 3
		dr.Type = 1
		dr.Open
		DoteyUpload_SourceData.position = FileStart
		DoteyUpload_SourceData.copyto dr, FileSize
		dr.SaveToFile MapPath(FullPath), 2
		dr.Close
		set dr = nothing 
		SaveAs = true
	End function

	'取得Server端路徑
	Private Function MapPath(Path)
		If InStr(1, Path, ":") > 0 Or Left(Path, 2) = "\\" Then
			MapPath = Path 
		Else 
			MapPath = Server.MapPath(Path)
		End If
	End function

	'根據路徑獲取文件名稱
	Private Function GetFileNameByPath(FullPath)
		Dim pos
		pos = 0
		FullPath = Replace(FullPath, "/", "\")
		pos = InStrRev(FullPath, "\") + 1
		If (pos > 0) Then
			GetFileNameByPath = Mid(FullPath, pos)
		Else
			GetFileNameByPath = FullPath
		End If
	End Function

End Class

%>
