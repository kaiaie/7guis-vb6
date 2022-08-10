# 7 GUIs in Visual Basic 6

This repository contains implementations of the [7GUIs exercises](https://eugenkiss.github.io/7guis/tasks/) 
in Microsoft Visual Basic 6. Because what good is a GUI challenge without an 
entry from one of the OG Rapid Application Development tools? ;-)


## Pre-requisites

You will require Visual Basic 6 to compile these exercises. VB6 has been out 
of support for nearly 20 years (though the VB runtime continues to be 
distributed with Windows). It is stil available as part of an MSDN 
subscriptions, though no doubt it can be downloaded from a warez site 
somewhere.


## Concluding thoughts

The following are a set of random thoughts having completed the exercises, 
both related to the exercises themselves and also how it was to implement them 
in a language that, as mentioned above, is largely obsolete.


### Thoughts on the exercises

I think that the exercises themselves are well thought-out and make for 
good tests of a front-end developer's skill; though some (like the "Cells" 
exercise) are probably too complex to be undertaken in an interview setting 
(it took me a little over a week, working on it in the evenings and at 
weekends). Having said that, I believe that even if the "Cells" exercise 
wasn't actually coded, the design itself would make for a good discussion of 
how an interview candidate might tackle it. There are a lot of interesting 
challenges in the exercise: how to track dependencies, for example, or how to 
write a simple parser.

One of the issues with the exercises is that they have to be pretty generic 
by their nature: they are intended to be completed by front-end developers 
using a variety of programming languages and UI toolkits and so cannot be tied 
to the conventions of a single environment. There's a few of the exercise 
where I feel there were compromises made in the interest of expediency, 
especially around issues like data validation, which were probably inevitable 
if developers were to have a chance of finishing them in a reasonable amount 
of time. However, it's often the case that the difference between a 
workmanlike GUI and one that users really enjoy interacting with is how well 
it conforms to the conventions of its host platform. A user interface that is 
a joy to use is often the result of a lot of tiny decisions that, 
individually, seem irrelevant but, taken together, make all the difference. 
One of the challenges I faced with these exercises was to try and resist as 
much as possible the temptation to keep polishing things, to make them as 
"Windows native" as possible (whatever that means in the modern world of 
fractured user experience on Windows, with flat "Metro"/ Microsoft Design 
Language/ Fluent applications, the "ribbon" UIs of Office, etc., but that is 
another rant) and to move onto the next exercise. It is one of the ironies of 
front-end development that time constraints often mean that a custom front-end 
does not get the same polish an off-the-shelf product does, yet an organisation 
is paying more for it!


### Thoughts on VB6

When the .NET Framework and the end of Visual Basic "Classic" was announced in 
the early 2000s, it made a lot of VB6 developers feel very unhappy and 
betrayed. The new C# language seemed to be the one getting all the press and 
Visual Basic.NET, for all the improvements like a much more conventional 
object model, felt like an afterthought in comparison. Much of the tooling 
supplied with the initial versions of .NET to migrate VB6 projects to the new 
platform wasn't very good, especially on the kind of complex line-of-business 
apps that VB6 developers wrote and maintained. And while VB.NET has sometimes 
gained new features before C# in the intervening years, it still feels very 
much that the latter language is where the action is.

(Aside: one of the big ironies of the .NET transition was how much of it was a 
wasted effort it was in retrospect; that is, it never really led to a 
Renaissance in the development of line-of-business applications that Microsoft 
might of hoped for. Despite my misgivings about the end of 
VB6, I remember being quite excited about what the .NET platform _promised_ 
for line-of-business developers like me. Proper threading for one, so that 
long-running business processes didn't block the user interface. But I don't 
think that promise was ever exploited to its fullest extent, mostly because 
front-end development was moving towards the Web. And a lot of the mistakes of 
the VB6 era ended up being replicated in the .NET era too: I've seen my share 
of .NET applications with thousands of lines of business logic embedded in 
event handlers, every bit as bad as any VB6 prototype application pushed 
into production. Perhaps there are teams of WinForms or WPF monks out there 
developing beautiful jewels of desktop applications using all the features of 
.NET that seemed so promising to me back then, but I haven't encountered them 
in nearly twenty years as a contractor. Make of that what you will.)

To step back into the world of VB6 after nearly two decades spent in .NET 
feels odd, for sure, but not completely terrible. One of the main advantages 
is how _fast_ the VB6 IDE is on modern hardware: it starts nearly instantly in 
the Windows XP virtual machine I used to develop the exercises, on a PC that is 
a good 7 years old. The VB6 IDE is also refreshingly _uncluttered_ compared to 
a modern version of Visual Studio with its endless gallery of tool windows, 
all extremely useful in their way, but somehow also always the one you _don't_ 
need right now.

I confess to often feeling a kind of "analysis paralysis" when developing in 
modern .NET and Visual Studio that the VS UI contributes to: there's a 
temptation to go heavy on the ceremony because this is a "professional" 
programming language and environment. In VB6, by comparison, I feel it's 
easier to jump straight in and iterate towards a solution, because no matter 
how elegant or ugly that solution turns out, the object-oriented orthodoxy 
will never deem it worthy by dint of it being written in VB6. In a sense, it 
makes VB6 _more_ of an agile language than its successor, at least to me!

Having sung VB6's praises, there is very much a feel of being in the world of 
"stone knives and bearskins" in some places with VB6. One thing I very much 
feel the lack of nowadays is a good (generic) collections library, and 
[LINQ](https://en.wikipedia.org/wiki/Language_Integrated_Query). I love LINQ 
and I love having collections one can iterate over with `For Each` without 
resorting to horrible COM black magic, and inline functions and much more. The 
COM `Collection` object that VB6 provides is a very poor substitute in 
comparison. I think that it is an expectation that modern languages come with 
a much more feature-rich runtime than was the standard in the 1990s and 2000s.
