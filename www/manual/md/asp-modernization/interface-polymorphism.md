# Interface Polymorphism (`Implements`)

AxonASP v2 supports VB6-style interface polymorphism using the `Implements` keyword. This allows a class to implement multiple interfaces and provide specific logic for each, which is essential for building modular and reusable components in Classic ASP.

## Overview

In VB6 and AxonASP, any class can act as an interface. When a class implements an interface, it must provide implementations for all public methods and properties of that interface using the naming convention `InterfaceName_MethodName`.

When an object is assigned to a variable explicitly typed as an interface (using the `As` clause), member calls on that variable are automatically routed to the interface-prefixed implementations.

## Usage

### Defining an Interface

An interface is simply a standard VBScript class with public method and property stubs.

```vbscript
' IAnimal.asp
Class IAnimal
    Function MakeSound()
    End Function
    
    Property Get Species()
    End Property
End Class
```

### Implementing the Interface

To implement an interface, use the `Implements` keyword inside the class definition. Implement each member with the `InterfaceName_` prefix.

```vbscript
Class Dog
    Implements IAnimal
    
    ' Implementation of IAnimal.MakeSound
    Function IAnimal_MakeSound()
        IAnimal_MakeSound = "Woof!"
    End Function
    
    ' Implementation of IAnimal.Species
    Property Get IAnimal_Species()
        IAnimal_Species = "Canine"
    End Property
    
    ' Default implementation (optional)
    Function MakeSound()
        MakeSound = "Generic Dog Sound"
    End Function
End Class
```

### Polymorphic Dispatch

When you type a variable as the interface, AxonASP routes calls to the interface implementations.

```vbscript
Dim myPet As IAnimal
Set myPet = New Dog

' This calls IAnimal_MakeSound on the Dog instance
Response.Write myPet.MakeSound() ' Output: Woof!

Dim untypedPet
Set untypedPet = New Dog

' This calls the default MakeSound on the Dog instance
Response.Write untypedPet.MakeSound() ' Output: Generic Dog Sound
```

## Key Features

1.  **Multiple Interfaces**: A single class can implement multiple interfaces.
2.  **Transparent Routing**: The VM handles method routing automatically based on the variable's declared type.
3.  **Encapsulation**: Interface implementations are typically `Private` in standard VBScript but remain accessible when called through the interface reference.
4.  **Zero Allocation**: Method resolution uses fast string-matching and cached lookups, maintaining high performance.

## Prerequisites

- AxonASP v2.0 or higher.
- `Option Explicit` is recommended but not required for interface polymorphism.
