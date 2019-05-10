Imports System
Imports System.ComponentModel
Imports System.Windows.Forms

Namespace FastColoredTextBoxNS
    Class FCTBDescriptionProvider
        Inherits TypeDescriptionProvider

        Public Sub New(ByVal type As Type)
            MyBase.New(GetDefaultTypeProvider(type))
        End Sub

        Private Shared Function GetDefaultTypeProvider(ByVal type As Type) As TypeDescriptionProvider
            Return TypeDescriptor.GetProvider(type)
        End Function

        Public Overrides Function GetTypeDescriptor(ByVal objectType As Type, ByVal instance As Object) As ICustomTypeDescriptor
            Dim defaultDescriptor As ICustomTypeDescriptor = MyBase.GetTypeDescriptor(objectType, instance)
            Return New FCTBTypeDescriptor(defaultDescriptor, instance)
        End Function
    End Class

    Class FCTBTypeDescriptor
        Inherits CustomTypeDescriptor

        Private parent As ICustomTypeDescriptor
        Private instance As Object

        Public Sub New(ByVal parent As ICustomTypeDescriptor, ByVal instance As Object)
            MyBase.New(parent)
            Me.parent = parent
            Me.instance = instance
        End Sub

        Public Overrides Function GetComponentName() As String
            Dim ctrl = (TryCast(instance, Control))
            Return If(ctrl Is Nothing, Nothing, ctrl.Name)
        End Function

        Public Overrides Function GetEvents() As EventDescriptorCollection
            Dim coll = MyBase.GetEvents()
            Dim list = New EventDescriptor(coll.Count - 1) {}

            For i As Integer = 0 To coll.Count - 1

                If coll(i).Name = "TextChanged" Then
                    list(i) = New FooTextChangedDescriptor(coll(i))
                Else
                    list(i) = coll(i)
                End If
            Next

            Return New EventDescriptorCollection(list)
        End Function
    End Class

    Class FooTextChangedDescriptor
        Inherits EventDescriptor

        Public Sub New(ByVal desc As MemberDescriptor)
            MyBase.New(desc)
        End Sub

        Public Overrides Sub AddEventHandler(ByVal component As Object, ByVal value As [Delegate])
            AddHandler(TryCast(component, FastColoredTextBox)).BindingTextChanged, TryCast(value, EventHandler)
        End Sub

        Public Overrides ReadOnly Property ComponentType As Type
            Get
                Return GetType(FastColoredTextBox)
            End Get
        End Property

        Public Overrides ReadOnly Property EventType As Type
            Get
                Return GetType(EventHandler)
            End Get
        End Property

        Public Overrides ReadOnly Property IsMulticast As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Sub RemoveEventHandler(ByVal component As Object, ByVal value As [Delegate])
            RemoveHandler(TryCast(component, FastColoredTextBox)).BindingTextChanged, TryCast(value, EventHandler)
        End Sub
    End Class
End Namespace