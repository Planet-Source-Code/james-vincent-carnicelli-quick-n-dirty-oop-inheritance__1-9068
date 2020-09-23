<div align="center">

## Quick N Dirty OOP Inheritance


</div>

### Description

VB (version 6 and earlier) doesn't offer any decent way to create a new class that inherits public members from an existing one. Here's an easy, if inelegant, way to simulate inheritance.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Vincent Carnicelli](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-vincent-carnicelli.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-vincent-carnicelli-quick-n-dirty-oop-inheritance__1-9068/archive/master.zip)





### Source Code

Most programmers who have any understanding of what object-oriented programming (OOP) is about have heard terms like "inheritance" and "subclassing". The goal is to create a new class starting not from scratch, but using an existing class as the foundation. While many other languages like C++ and Java offer inheritance models, Visual Basic 6 and earlier versions don't in any decent sense. The closest it comes to it is the use of the messy "Implements" directive.
<P>What most programmers familiar with OOP don’t know is that there are two basic relationships with which to implement inheritance: "is a" and "has a". Let's say for example we have the following three classes: Animal, Dog, and Beagle. We want Dog to inherit public members from Animal and Beagle to inherit them from Dog. Speaking "purely" of OOP, we would say that we would say that a Beagle <B>is a</B> Dog. The alternative would be to say that a Beagle <B>has a</B> Dog. In English, this sounds like nonsense, but bear with me. If you want to gain the functionality of one class, it suffices to simply instantiate it, which is another way of saying the first class would <B>have an</B> instance of the other. Consider the following illustration:
<CENTER>
<TABLE BGCOLOR="#FFFFCC" CELLSPACING="0" CELLPADDING="4" BORDER="1"><TR><TD>
<CENTER><B>Beagle</B></CENTER>
<CENTER>
<TABLE><TR><TD VALIGN="TOP"><NOBR><LI>BaseClass As </NOBR></TD><TD>
<TABLE BGCOLOR="#EEEEBB" CELLSPACING="0" CELLPADDING="4" BORDER="1"><TR><TD>
<CENTER><B>Dog</B></CENTER>
<CENTER>
<TABLE><TR><TD VALIGN="TOP"><LI><NOBR>BaseClass As </NOBR></TD><TD>
<TABLE BGCOLOR="#DDDDAA" CELLSPACING="0" CELLPADDING="4" BORDER="1"><TR><TD>
<CENTER><B>Animal</B></CENTER>
<NOBR>
<LI>Species As String
</NOBR>
</TD></TR></TABLE>
</TD></TR></TABLE>
</CENTER>
<LI>HasFleas As Boolean
</TD></TR></TABLE>
</TD></TR></TABLE>
</CENTER>
<LI>HasLongEars As Boolean
<LI>HasFleas As Boolean
</TD></TR></TABLE>
</CENTER>
<P>Note the "overloaded" <TT>.HasFleas</TT> property in both Dog and Beagle. Here are the equivalent VB class modules:
<CENTER>
<P><TABLE BGCOLOR="#FFFFCC" WIDTH="90%" CELLSPACING="0" CELLPADDING="4" BORDER="1"><TR><TH BGCOLOR="#DDDDAA"> Class: Animal </TH></TR><TR><TD><PRE>
<FONT COLOR="#000099">Private</FONT> propSpecies <FONT COLOR="#000099">As String</FONT>
<BR>
<BR><FONT COLOR="#000099">Public Property Get</FONT> Species() <FONT COLOR="#000099">As String</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;Species = propSpecies
<FONT COLOR="#000099">End Property</FONT>
<FONT COLOR="#000099">Public Property Let</FONT> Species(newSpecies <FONT COLOR="#000099">As String</FONT>)
&nbsp;&nbsp;&nbsp;&nbsp;propSpecies = newSpecies
<FONT COLOR="#000099">End Property</FONT>
</PRE></TD></TR></TABLE>
<P><TABLE BGCOLOR="#FFFFCC" WIDTH="90%" CELLSPACING="0" CELLPADDING="4" BORDER="1"><TR><TH BGCOLOR="#DDDDAA"> Class: Dog </TH></TR><TR><TD><PRE>
<FONT COLOR="#000099">Private</FONT> <FONT COLOR="#CC0000"><B>BaseClass</B></FONT> <FONT COLOR="#000099">As Animal</FONT>
<FONT COLOR="#000099">Private</FONT> propHasFleas <FONT COLOR="#000099">As Boolean</FONT>
<BR>
<BR><FONT COLOR="#000099">Public Property Get</FONT> <FONT COLOR="#CC0000"><B>B()</B></FONT> <FONT COLOR="#000099">As Animal</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Set</FONT> B = BaseClass
<FONT COLOR="#000099">End Property</FONT>
<BR>
<BR><FONT COLOR="#000099">Public Property Get</FONT> HasFleas() <FONT COLOR="#000099">As Boolean</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;HasFleas = propHasFleas
<FONT COLOR="#000099">End Property</FONT>
<FONT COLOR="#000099">Public Property Let</FONT> HasFleas(newHasFleas <FONT COLOR="#000099">As Boolean</FONT>)
&nbsp;&nbsp;&nbsp;&nbsp;propHasFleas = newHasFleas
<FONT COLOR="#000099">End Property</FONT>
<BR>
<BR><FONT COLOR="#000099">Private Sub</FONT> Class_Initialize()
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#CC0000"><B>Set BaseClass = New Animal</B></FONT>
&nbsp;&nbsp;&nbsp;&nbsp;BaseClass.Species = "Canus"
<FONT COLOR="#000099">End Sub</FONT>
</PRE></TD></TR></TABLE>
<P><TABLE BGCOLOR="#FFFFCC" WIDTH="90%" CELLSPACING="0" CELLPADDING="4" BORDER="1"><TR><TH BGCOLOR="#DDDDAA"> Class: Beagle </TH></TR><TR><TD><PRE>
<FONT COLOR="#000099">Private</FONT> <FONT COLOR="#CC0000"><B>BaseClass</B></FONT> <FONT COLOR="#000099">As Dog</FONT>
<FONT COLOR="#000099">Private</FONT> propHasFleas <FONT COLOR="#000099">As Boolean</FONT>
<FONT COLOR="#000099">Private</FONT> propHasLongEars <FONT COLOR="#000099">As Boolean</FONT>
<BR>
<BR><FONT COLOR="#000099">Public Property Get</FONT> <FONT COLOR="#CC0000"><B>B()</B></FONT> <FONT COLOR="#000099">As Dog</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#000099">Set</FONT> B = BaseClass
<FONT COLOR="#000099">End Property</FONT>
<BR>
<BR><FONT COLOR="#000099">Public Property Get</FONT> HasFleas() <FONT COLOR="#000099">As Boolean</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;HasFleas = <FONT COLOR="#000099">True</FONT>
<FONT COLOR="#000099">End Property</FONT>
<BR>
<BR><FONT COLOR="#000099">Public Property Get</FONT> HasLongEars() <FONT COLOR="#000099">As Boolean</FONT>
&nbsp;&nbsp;&nbsp;&nbsp;HasLongEars = <FONT COLOR="#000099">True</FONT>
<FONT COLOR="#000099">End Property</FONT>
<BR>
<BR><FONT COLOR="#000099">Private Sub</FONT> Class_Initialize()
&nbsp;&nbsp;&nbsp;&nbsp;<FONT COLOR="#CC0000"><B>Set BaseClass = New Dog</B></FONT>
<FONT COLOR="#000099">End Sub</FONT>
</PRE></TD></TR></TABLE>
</CENTER>
<P>So when we create a Beagle object, it's creating a Dog object internally, which in tun is creating an Animal object inside itself. So if Beagle "inherits" functionality from Dog and Dog likewise from Animal, how to we use this inherited functionality in our code? One answer which is elegant from the Beagle class's user's (programmer's) point of view would be to reproduce properties, methods, and events with the same names as all of what's being "inherited". So Beagle, for example, would have a <TT>.Species</TT> property to mirror the one in Animal which would simply delegate the work of storage and/or processing to the Animal class. But then, the chore of creating these mirror members can really suck if you're making a class that adds only two new members to a class that already has two dozen you want to inherit.
<P>A simpler way, which is admittedly a little messier for the end programmer, is to give him a handle to the "base class" object so he can directly access its members. We've done this by adding a <TT>.B</TT> -- short for "Base Class" -- property. So to find out what species our beagle is we might say <TT>TheSpecies = MyBeagle.B.B.Species</TT>. Granted, this isn't as elegant as <TT>TheSpecies = MyBeagle.Species</TT>, but this construct is invalid in our case. Suffice it to say that using the slighly ugly <TT>.B</TT> property to get to an object's "base class" works and doesn't take much effort on your part to implement. It's also worth pointing out that you can do multiple inheritance this way, too, so long as you come up with a different property name for each of the base classes. You might use <TT>.B1</TT> and <TT>.B2</TT> or perhaps opt to go with more explicitly named properties like <TT>.BcDog</TT> and <TT>.BcAnimal</TT>.
<P>The key to making this work for the user, who most likely won't care how your class is implemented, is to instruct him to look for a given property or method in your object, first, and then to look for a property or method in the object <TT>.B</TT> refers to if he can't find it in your class directly. This is especially important if you overload a given function, as in our case where <TT>.HasFleas</TT> is implemented in the Dog class but also in the Beagle class. The user of the Beagle class can refer, then, to <TT>.HasFleas</TT> to get your overriding property or to the <TT>.B.HasFleas</TT> property to refer to the overridden version of the property.
<P>While this mildly messy of approach may not be an elegant one if you're distributing polished products to clients or trying to set industry standards, it's an excellent way to help organize and maintain the inner workings of your more complicated VB projects.
<P>In summary, although VB doesn't have a clean implementation of the traditional OOP inheritance ("is a") concept, you can simulate it using a "has a" relationship. The syntax for using the "inherited" members may seem a bit awkward, but the benefit is a simpler implementation for you and a greater ease of maintaining your code, both encouraging you to better modularize your code toward the ends modularization has long promised.

