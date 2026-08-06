using JetBrains.Annotations;

namespace Parchment;

/// <summary>
/// Declares that the decorated <c>partial</c> class is a Parchment binding model whose member
/// accessors are pre-compiled by the source generator — without binding a template. For models
/// registered by hand via <see cref="TemplateStore.RegisterDocxTemplate{TModel}(Stream, ProtectionMode)"/> /
/// <see cref="TemplateStore.RegisterMarkdownTemplate{TModel}"/> against templates the generator
/// cannot see: content produced at runtime, per-tenant templates, and the like.
///
/// A model bound to a compile-time template uses <see cref="ParchmentModelAttribute"/> instead,
/// which embeds the template as well; a class never needs both.
///
/// The <c>[MeansImplicitUse]</c> annotation tells ReSharper / Rider that members of the
/// decorated class are bound implicitly at render time, so it stops emitting
/// "Property is never used" / <c>UnusedAutoPropertyAccessor.Global</c> warnings on them.
/// </summary>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
[MeansImplicitUse(ImplicitUseTargetFlags.Members)]
public sealed class ParchmentBindableAttribute :
    Attribute;
