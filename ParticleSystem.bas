Attribute VB_Name = "ParticleSystem"
' Primitive Particle System v1.0


Public Type Particle
    X As Single         ' X-Pos of particle
    Y As Single         ' Y-Pos of particle
    VX As Single        ' Velocity along X-axis
    VY As Single        ' Velocity along Y-axis
    Color As Long       ' Color, can mean many things according to the specific program
    Size As Single      ' Size of particle, can also mean many things
    Age As Long         ' Age of particle, as some particles "die out" as their energy drops.
End Type

' Array to hold particles (may be upgraded to linked list in upcoming version)
Public ParticleSys() As Particle

' CreateParticle - Takes a Particle type and adds it to the ParticleSys array, returning
' the index of the particle
Public Function CreateParticle(TheParticle As Particle) As Long
    On Error Resume Next
    Dim l As Long
    
    l = UBound(ParticleSys)
    ReDim Preserve ParticleSys(l + 1)
    With ParticleSys(l + 1)
        .X = TheParticle.X
        .Y = TheParticle.Y
        .VX = TheParticle.VX
        .VY = TheParticle.VY
        .Color = TheParticle.Color
        .Size = TheParticle.Size
        .Age = 1
    End With
    CreateParticle = l + 1
End Function

' DeleteParticle - Removes a particle from the ParticleSys array
Public Function DeleteParticle(Index As Long) As Long
    On Error Resume Next
    Dim l As Long
    
    l = UBound(ParticleSys)
    With ParticleSys(Index)
        .X = ParticleSys(l).X
        .Y = ParticleSys(l).Y
        .VX = ParticleSys(l).VX
        .VY = ParticleSys(l).VY
        .Color = ParticleSys(l).Color
        .Size = ParticleSys(l).Size
    End With
    ReDim Preserve ParticleSys(l - 1)
    DeleteParticle = l
End Function

' UpdParticles - Updates POSITION only (velocity changing can be implemented in different ways
' so I decided that it was better to leave it to the specific program to handle velocity)
Public Sub UpdParticles()
    On Error Resume Next
    Dim i As Long
    Dim l As Long
    
    l = UBound(ParticleSys)
    For i = 1 To l
        With ParticleSys(i)
            .X = .X + .VX
            .Y = .Y + .VY
        End With
    Next i
End Sub

