### Initial project of a7VisualStudioExtensions

For now this extension just adds a `Replace initializer with constructor` command to the editor context menu.

#### Example:

When we select the object initializer part of below class instance creation:

```
var s = new SomeClass
{
	SomeField = "some value",
	SomeOtherField = "some other value"
};
```

and select the `Replace initializer with constructor` from the context menu, or press `Ctrl+2`, it will be replaced with:

```
var s = new SomeClass
(
	someField: "some value",
	someOtherField: "some other value"
);
```