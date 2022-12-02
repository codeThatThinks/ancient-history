Public Structure player
    Public x As Integer
    Public y As Integer
    Public name As String
End Structure

Public Structure item
    Public x As Integer
    Public y As Integer
    Public type As Integer
End Structure

Public Class frmMain
    Public blockFloor As Image, blockWall As Image, blockStart As Image, texturePlayer As Image, _
    itemEmerald As Image, itemRuby As Image, itemDiamond As Image, itemGold As Image, itemRope As Image, _
    itemSword As Image, itemTrapdoor As Image, itemDragon As Image

    Public myPlayer As player, pointStart As Point

    Public walls(0) As Point, items(7) As item, inv(7) As String

    Public isDebug As Boolean, currentMap As String, isTrapped As Boolean

    Public Sub hideItems()
        loadMap(currentMap)
    End Sub

    Public Sub showItems()
        Dim imgBackground As Image
        Dim g As Graphics

        loadMap(currentMap)

        imgBackground = pictCanvas.Image
        g = Graphics.FromImage(imgBackground)

        ' 0 = emerald
        ' 1 = ruby
        ' 2 = diamond
        ' 3 = gold
        ' 4 = rope
        ' 5 = sword
        ' 6 = trap door
        ' 7 = dragon's lair

        For i As Integer = 0 To items.GetUpperBound(0)
            Select Case items(i).type
                Case 0
                    g.DrawImageUnscaled(itemEmerald, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 1
                    g.DrawImageUnscaled(itemRuby, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 2
                    g.DrawImageUnscaled(itemDiamond, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 3
                    g.DrawImageUnscaled(itemGold, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 4
                    g.DrawImageUnscaled(itemRope, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 5
                    g.DrawImageUnscaled(itemSword, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 6
                    g.DrawImageUnscaled(itemTrapdoor, items(i).x * 42, items(i).y * 42)
                    Exit Select

                Case 7
                    g.DrawImageUnscaled(itemDragon, items(i).x * 42, items(i).y * 42)
                    Exit Select
            End Select
        Next

        pictCanvas.Image = imgBackground
        g.Dispose()
    End Sub

    Public Function findInv(ByVal type As String)
        Dim i As Integer
        For i = 0 To inv.GetUpperBound(0)
            If inv(i) = type Then
                Return True
                Exit For
            End If
        Next

        Return False
    End Function

    Public Sub addInv(ByVal type As String)
        Dim i As Integer, invBox As Windows.Forms.PictureBox
        For i = 0 To inv.GetUpperBound(0)
            If inv(i) = "" Then
                inv(i) = type
                Exit For
            End If
        Next

        If isDebug Then lblChat.Text = "New inv(" & i & ") = " & type

        Select Case i
            Case 0
                invBox = pictInv1
                Exit Select

            Case 1
                invBox = pictInv2
                Exit Select

            Case 2
                invBox = pictInv3
                Exit Select

            Case 3
                invBox = pictInv4
                Exit Select

            Case 4
                invBox = pictInv5
                Exit Select

            Case 5
                invBox = pictInv6
                Exit Select
        End Select

        Select Case type
            Case "emerald"
                invBox.Image = itemEmerald
                Exit Select

            Case "ruby"
                invBox.Image = itemRuby
                Exit Select

            Case "diamond"
                invBox.Image = itemDiamond
                Exit Select

            Case "gold"
                invBox.Image = itemGold
                Exit Select

            Case "rope"
                invBox.Image = itemRope
                Exit Select

            Case "sword"
                invBox.Image = itemSword
                Exit Select
        End Select
    End Sub

    Public Sub removeItem(ByVal x As Integer, ByVal y As Integer)
        For i As Integer = 0 To items.GetUpperBound(0)
            If items(i).x = x And items(i).y = y Then
                items(i).type = 10                          ' set to bogus type to remove it
                Exit For
            End If
        Next

        If isDebug Then showItems() ' refresh items on screen
    End Sub

    Public Function getItemType(ByVal x As Integer, ByVal y As Integer)
        For i As Integer = 0 To items.GetUpperBound(0)      ' search items array
            If items(i).x = x And items(i).y = y Then
                Return items(i).type                        ' return item type
                Exit For
            End If
        Next
        Return False                                        ' return false if not found
    End Function

    Public Function isItem(ByVal x As Integer, ByVal y As Integer)
        For i As Integer = 0 To items.GetUpperBound(0)      ' search items array
            If items(i).x = x And items(i).y = y And Not items(i).type = 10 Then
                Return True                                 ' return true if found
                Exit For
            End If
        Next
        Return False                                        ' return false if not found
    End Function

    Public Function isWall(ByVal x As Integer, ByVal y As Integer)
        For i As Integer = 0 To walls.GetUpperBound(0)      ' search walls array
            If walls(i) = New Point(x, y) Then
                Return True                                 ' return true if found
                Exit For
            End If
        Next
        Return False                                        ' return false if not found
    End Function

    Public Sub drawPlayer()
        pictPlayer.Image = texturePlayer
        pictPlayer.Left = myPlayer.x * 42 + 12
        pictPlayer.Top = myPlayer.y * 42 + 12
        pictPlayer.Show()
    End Sub

    Public Sub movePlayer(ByVal direction As String)
        Dim newCoordsX As Integer, newCoordsY As Integer

        lblChat.Text = ""

        Select Case direction
            Case "up"
                newCoordsX = myPlayer.x
                newCoordsY = myPlayer.y - 1

            Case "down"
                newCoordsX = myPlayer.x
                newCoordsY = myPlayer.y + 1

            Case "left"
                newCoordsX = myPlayer.x - 1
                newCoordsY = myPlayer.y

            Case "right"
                newCoordsX = myPlayer.x + 1
                newCoordsY = myPlayer.y
        End Select

        If Not isWall(newCoordsX, newCoordsY) Then
            myPlayer.x = newCoordsX
            myPlayer.y = newCoordsY
        End If

        drawPlayer()

        ' items types:
        ' 0 = emerald
        ' 1 = ruby
        ' 2 = diamond
        ' 3 = gold
        ' 4 = rope
        ' 5 = sword
        ' 6 = trap door
        ' 7 = dragon's lair

        If isDebug Then lblChat.Text = "isItem() = " & isItem(myPlayer.x, myPlayer.y) & "    getItemType() = " & getItemType(myPlayer.x, myPlayer.y)

        If isItem(myPlayer.x, myPlayer.y) Then
            Select Case getItemType(myPlayer.x, myPlayer.y)
                Case 0
                    If Not isDebug Then lblChat.Text = "You have found emerald."
                    addInv("emerald")
                    removeItem(myPlayer.x, myPlayer.y)
                    Exit Select

                Case 1
                    If Not isDebug Then lblChat.Text = "You have found ruby."
                    addInv("ruby")
                    removeItem(myPlayer.x, myPlayer.y)
                    Exit Select

                Case 2
                    If Not isDebug Then lblChat.Text = "You have found diamond."
                    addInv("diamond")
                    removeItem(myPlayer.x, myPlayer.y)
                    Exit Select

                Case 3
                    If Not isDebug Then lblChat.Text = "You have found gold."
                    addInv("gold")
                    removeItem(myPlayer.x, myPlayer.y)
                    Exit Select

                Case 4
                    If Not isDebug Then lblChat.Text = "You have found rope. Press <space> to use it."
                    addInv("rope")
                    removeItem(myPlayer.x, myPlayer.y)
                    Exit Select

                Case 5
                    If Not isDebug Then lblChat.Text = "You have found a sword. Press <ctrl> to use it."
                    addInv("sword")
                    removeItem(myPlayer.x, myPlayer.y)
                    Exit Select

                Case 6
                    isTrapped = True
                    currentMap = "trapped.map"
                    loadMap(currentMap)

                    If Not isDebug And findInv("rope") Then lblChat.Text = "You fell through a trap door. Press <space> to use your rope."
                    If Not findInv("rope") Then
                        MsgBox("You fell though a trapdoor without a rope and died.")
                        Me.Close()
                    End If
                    
                    Exit Select

                Case 7
                    currentMap = "dragon.map"
                    loadMap(currentMap)           ' load new dragon map

                    myPlayer.x = pointStart.X     ' set player start position
                    myPlayer.y = pointStart.Y

                    drawPlayer()                    ' draw the player
                    Array.Resize(items, 0)          ' clear items array

                    If Not isDebug Then lblChat.Text = "You have found a dragon's lair."
                    Exit Select
            End Select
        End If
    End Sub

    Public Sub createItems()
        Dim x As Integer, y As Integer, itemType As Integer

        Randomize()

        ' generate treasures
        For i As Integer = 0 To 3   ' 4 treasures total
            Do
                x = Int(Rnd() * 9 + 0)
                y = Int(Rnd() * 9 + 0)
            Loop Until Not isWall(x, y) And Not isItem(x, y) And Not pointStart.X = x And Not pointStart.Y = y

            itemType = Int(Rnd() * 3)

            ' add items coords to items array
            items(i).x = x
            items(i).y = y
            items(i).type = itemType
        Next

        ' generate rope
        Do
            x = Int(Rnd() * 9 + 0)
            y = Int(Rnd() * 9 + 0)
        Loop Until Not isWall(x, y) And Not isItem(x, y) And Not pointStart.X = x And Not pointStart.Y = y
        items(4).x = x
        items(4).y = y
        items(4).type = 4

        'generate sword
        Do
            x = Int(Rnd() * 9 + 0)
            y = Int(Rnd() * 9 + 0)
        Loop Until Not isWall(x, y) And Not isItem(x, y) And Not pointStart.X = x And Not pointStart.Y = y
        items(5).x = x
        items(5).y = y
        items(5).type = 5

        ' generate trapdoor
        Do
            x = Int(Rnd() * 9 + 0)
            y = Int(Rnd() * 9 + 0)
        Loop Until Not isWall(x, y) And Not isItem(x, y) And Not pointStart.X = x And Not pointStart.Y = y
        items(6).x = x
        items(6).y = y
        items(6).type = 6

        ' generate dragon 's lair
        Do
            x = Int(Rnd() * 9 + 0)
            y = Int(Rnd() * 9 + 0)
        Loop Until Not isWall(x, y) And Not isItem(x, y) And Not pointStart.X = x And Not pointStart.Y = y
        items(7).x = x
        items(7).y = y
        items(7).type = 7
    End Sub

    Public Sub loadMap(ByVal mapPath As String)
        Dim imgBackground As New Bitmap(420, 420)
        Dim g As Graphics = Graphics.FromImage(imgBackground)

        Dim mapFile As IO.StreamReader = New IO.StreamReader(mapPath)
        Dim mapData As String, mapChunk As String
        Dim mapRenderRow As Integer, mapRenderColumn As Integer

        Array.Resize(walls, 0)      ' clear walls array

        Do Until mapFile.EndOfStream
            mapData = mapFile.ReadLine()

            For mapRenderColumn = 0 To mapData.Length - 1
                mapChunk = mapData.Substring(mapRenderColumn, 1)

                ' 0 = open space, default floor texture
                ' X = wall
                ' S = starting position, looks like door
                ' D = dragon

                Select Case mapChunk
                    Case "O"
                        ' draw block
                        g.DrawImage(blockFloor, (mapRenderColumn) * 42, (mapRenderRow) * 42, 42, 42)
                        Exit Select

                    Case "D"
                        ' draw block
                        g.DrawImage(blockFloor, (mapRenderColumn) * 42, (mapRenderRow) * 42, 42, 42)
                        Exit Select

                    Case "S"
                        ' draw block
                        g.DrawImage(blockStart, (mapRenderColumn) * 42, (mapRenderRow) * 42, 42, 42)

                        ' set starting position to block coords
                        pointStart.X = mapRenderColumn
                        pointStart.Y = mapRenderRow
                        Exit Select

                    Case "X"
                        ' draw block
                        g.DrawImage(blockWall, (mapRenderColumn) * 42, (mapRenderRow) * 42, 42, 42)

                        ' add block coords to walls array
                        Array.Resize(walls, walls.Length + 1)
                        walls(walls.GetUpperBound(0)).X = mapRenderColumn
                        walls(walls.GetUpperBound(0)).Y = mapRenderRow
                        Exit Select

                End Select
                pictCanvas.Image = imgBackground
            Next

            mapRenderRow += 1
        Loop

        If isTrapped Then
            g.DrawImage(blockStart, myPlayer.x * 42, myPlayer.y * 42, 42, 42)
            pictCanvas.Image = imgBackground
        End If

        pictCanvas.Focus()

        ' cleanup
        g.Dispose()
        mapFile.Close()
    End Sub

    Public Sub loadTextures()
        blockFloor = Image.FromFile("textures\blocks\floor.png")
        blockWall = Image.FromFile("textures\blocks\wall.png")
        blockStart = Image.FromFile("textures\blocks\start.png")
        texturePlayer = Image.FromFile("textures\player.png")
        itemEmerald = Image.FromFile("textures\items\emerald.png")
        itemRuby = Image.FromFile("textures\items\ruby.png")
        itemDiamond = Image.FromFile("textures\items\diamond.png")
        itemGold = Image.FromFile("textures\items\gold.png")
        itemRope = Image.FromFile("textures\items\rope.png")
        itemSword = Image.FromFile("textures\items\sword.png")
        itemTrapdoor = Image.FromFile("textures\items\trapdoor.png")
        itemDragon = Image.FromFile("textures\items\dragon.png")
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ' dispose of images
        blockFloor.Dispose()
        blockWall.Dispose()
        blockStart.Dispose()
        texturePlayer.Dispose()
        itemEmerald.Dispose()
        itemRuby.Dispose()
        itemDiamond.Dispose()
        itemGold.Dispose()
        itemRope.Dispose()
        itemSword.Dispose()
        itemTrapdoor.Dispose()
        itemDragon.Dispose()
    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Up Or e.KeyCode = Keys.W Then movePlayer("up")
        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.S Then movePlayer("down")
        If e.KeyCode = Keys.Left Or e.KeyCode = Keys.A Then movePlayer("left")
        If e.KeyCode = Keys.Right Or e.KeyCode = Keys.D Then movePlayer("right")

        If e.KeyCode = Keys.I Then
            If isDebug Then
                isDebug = False
                hideItems()
                lblChat.Text = "Debug mode disabled."
            Else
                isDebug = True
                If Not isTrapped Then showItems()
                lblChat.Text = "Debug mode enabled."
            End If
        End If

        If e.KeyCode = Keys.Space And isTrapped Then
            If findInv("rope") Then
                isTrapped = False
                currentMap = "default.map"
                loadMap(currentMap)

                If isDebug Then showItems()
                If Not isDebug Then lblChat.Text = "You have successfully rescued yourself."
            End If
        End If
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        loadTextures()                  ' load textures

        currentMap = "default.map"
        loadMap(currentMap)             ' load map
        createItems()                   ' place items

        frmHowTo.ShowDialog()           ' show how to play form

        myPlayer.x = pointStart.X       ' set player start position
        myPlayer.y = pointStart.Y

        drawPlayer()                    ' draw the player

        lblChat.Text = "Hello Player!"
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim g As Graphics = pictPlayer.CreateGraphics

        g.DrawImage(texturePlayer, 0, 0, 42, 42)
    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmHowTo.ShowDialog()
    End Sub
End Class