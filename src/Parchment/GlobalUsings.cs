global using System;
global using System.Collections.Concurrent;
global using System.Collections.Generic;
global using System.Diagnostics.CodeAnalysis;
global using System.Globalization;
global using System.IO;
global using System.Linq;
global using System.Reflection;
global using System.Text;
global using System.Threading;
global using System.Threading.Tasks;
global using System.Xml;
global using System.Xml.Linq;

global using DocumentFormat.OpenXml;
global using DocumentFormat.OpenXml.Packaging;
global using DocumentFormat.OpenXml.Wordprocessing;
global using Excelsior;
global using Fluid;
global using Fluid.Accessors;
global using Fluid.Ast;
global using Fluid.Values;

global using Markdig;
global using Markdig.Renderers;
global using Markdig.Renderers.Html;
global using Markdig.Syntax;
global using Markdig.Syntax.Inlines;

global using Microsoft.Extensions.Logging;
global using Microsoft.Extensions.Logging.Abstractions;
global using OpenXmlHtml;
global using Parchment;
global using Parchment.Generated;

// Excelsior exposes a DocumentProperties of its own, and this project imports both namespaces.
// Parchment's is the one meant here; Excelsior's is reachable as Excelsior.DocumentProperties.
global using DocumentProperties = Parchment.DocumentProperties;
