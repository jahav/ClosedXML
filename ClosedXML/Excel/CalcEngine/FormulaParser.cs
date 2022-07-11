using ClosedXML.Excel.CalcEngine.Exceptions;
using Irony.Ast;
using Irony.Parsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLParser;

namespace ClosedXML.Excel.CalcEngine
{
    /// <summary>
    /// A parser that takes a string and parses it into concrete syntax tree through XLParser and then
    /// to abstract syntax tree that is used to evaluate the formula.
    /// </summary>
    internal class FormulaParser
    {
        private const string defaultFunctionNameSpace = "_xlfn";
        // ReferenceItem term is a transient, so we also have to enumerate all terms in the ReferenceItem rule.
        private static readonly HashSet<string> ReferenceItemTerms = new HashSet<string>() { GrammarNames.Cell, GrammarNames.NamedRange, GrammarNames.VerticalRange, GrammarNames.HorizontalRange, GrammarNames.RefError, GrammarNames.UDFunctionCall, GrammarNames.StructuredReference };
        private static readonly IDictionary<string, ErrorExpression.ExpressionErrorType> ErrorMap = new Dictionary<string, ErrorExpression.ExpressionErrorType>(StringComparer.OrdinalIgnoreCase)
        {
            ["#REF!"] = ErrorExpression.ExpressionErrorType.CellReference,
            ["#VALUE!"] = ErrorExpression.ExpressionErrorType.CellValue,
            ["#DIV/0!"] = ErrorExpression.ExpressionErrorType.DivisionByZero,
            ["#NAME?"] = ErrorExpression.ExpressionErrorType.NameNotRecognized,
            ["#N/A"] = ErrorExpression.ExpressionErrorType.NoValueAvailable,
            ["#NULL!"] = ErrorExpression.ExpressionErrorType.NullValue,
            ["#NUM!"] = ErrorExpression.ExpressionErrorType.NumberInvalid
        };

        // TODO: Remove later, we only need GetExternalObject method, extract it here.
        private readonly CalcEngine _engine;
        private readonly Dictionary<string, FunctionDefinition> _fnTbl; // table with constants and functions (pi, sin, etc)
        private Dictionary<BnfTerm, BinaryOp> _binaryOpMap;
        private readonly Parser _parser;

        public FormulaParser(CalcEngine engine, Dictionary<string, FunctionDefinition> fnTbl)
        {
            _engine = engine;
            var grammar = GetGrammar();
            _binaryOpMap = new Dictionary<BnfTerm, BinaryOp> {
                { grammar.expop, BinaryOp.Exp },
                { grammar.mulop, BinaryOp.Mult },
                { grammar.divop, BinaryOp.Div },
                { grammar.plusop, BinaryOp.Add },
                { grammar.minop, BinaryOp.Sub },
                { grammar.concatop, BinaryOp.Concat},
                { grammar.gtop, BinaryOp.Gt},
                { grammar.eqop, BinaryOp.Eq },
                { grammar.ltop, BinaryOp.Lt },
                { grammar.neqop, BinaryOp.Neq },
                { grammar.gteop, BinaryOp.Gte },
                { grammar.lteop, BinaryOp.Lte },
            };
            _parser = new Parser(grammar);
            _fnTbl = fnTbl;
        }

        public Expression ParseToAst(string formula)
        {
            try
            {
                var tree = _parser.Parse(formula);
                return (Expression)tree.Root.AstNode;
            }
            catch (NullReferenceException ex) when (ex.StackTrace.StartsWith("   at Irony.Ast.AstBuilder.BuildAst(ParseTreeNode parseNode)"))
            {
                throw new InvalidProgramException($"Unable to parse formula '{formula}'. Some Irony grammar term is missing AST configuration.");
            }
        }

        private ExcelFormulaGrammar GetGrammar()
        {
            // Keep AST configuration in same order as is the 'SomeTerm.Rule ='  in in ExcelFormulaGrammar for readability.
            var grammar = new ExcelFormulaGrammar();
            grammar.FormulaWithEq.AstConfig.NodeCreator = CreateCopyNode(1);
            grammar.Formula.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.ArrayFormula.AstConfig.NodeCreator = CreateNotSupportedNode("array formula");

            grammar.MultiRangeFormula.AstConfig.NodeCreator = CreateCopyNode(1);
            grammar.Union.AstConfig.NodeCreator = CreateNotSupportedNode("range union operator");
            grammar.intersectop.AstConfig.NodeCreator = DontCreateNode;

            grammar.Constant.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.Number.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.NumberToken.AstConfig.NodeCreator = CreateNumberNode;
            grammar.Error.AstConfig.NodeCreator = CreateErrorNode;
            grammar.ErrorToken.AstConfig.NodeCreator = DontCreateNode;

            // RefErrorToken is marked with NoAstToken
            grammar.RefError.AstConfig.NodeCreator = CreateErrorNode;
            grammar.RefErrorToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.ConstantArray.AstConfig.NodeCreator = CreateNotSupportedNode("constant array");
            grammar.ArrayColumns.AstConfig.NodeCreator = DontCreateNode;
            grammar.ArrayRows.AstConfig.NodeCreator = DontCreateNode;
            grammar.ArrayConstant.AstConfig.NodeCreator = DontCreateNode;

            grammar.FunctionCall.AstConfig.NodeCreator = CreateFunctionCallNode;
            grammar.Argument.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.FunctionName.AstConfig.NodeCreator = DontCreateNode;
            grammar.ExcelFunction.AstConfig.NodeCreator = DontCreateNode;

            grammar.Arguments.AstConfig.NodeCreator = DontCreateNode;
            grammar.EmptyArgument.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.EmptyArgumentToken.AstConfig.NodeCreator = CreateEmptyArgumentNode;

            grammar.Bool.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.BoolToken.AstConfig.NodeCreator = CreateBoolNode;

            grammar.Text.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.TextToken.AstConfig.NodeCreator = CreateTextNode;

            // TODO: this is placeholder
            grammar.Reference.AstConfig.NodeCreator = CreateReferenceNode;
            grammar.Cell.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.CellToken.AstConfig.NodeCreator = CreateCellNode;
            grammar.NamedRange.AstConfig.NodeCreator = CreateNamedRangeNode;
            grammar.NameToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.NamedRangeCombinationToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.ReferenceFunctionCall.AstConfig.NodeCreator = CreateReferenceFunctionCallNode;
            grammar.RefFunctionName.AstConfig.NodeCreator = DontCreateNode;
            grammar.ExcelConditionalRefFunctionToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.ExcelRefFunctionToken.AstConfig.NodeCreator = DontCreateNode;

            // Prefix is only used in Reference term together with ReferenceItem. It is taken care of in CreateReferenceFunctionCallNode.
            grammar.Prefix.AstConfig.NodeCreator = DontCreateNode;
            grammar.SheetToken.AstConfig.NodeCreator = DontCreateNode;
            grammar.SheetQuotedToken.AstConfig.NodeCreator = DontCreateNode;

            // DDE formula parsing in XLParser seems to be buggy. It can't parse few examples I have found.
            grammar.DynamicDataExchange.AstConfig.NodeCreator = CreateNotSupportedNode("dynamic data exchange");
            grammar.SingleQuotedStringToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.VRange.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.VRangeToken.AstConfig.NodeCreator = CreateVerticalOrHorizontalRangeNode;

            grammar.HRange.AstConfig.NodeCreator = CreateCopyNode(0);
            grammar.HRangeToken.AstConfig.NodeCreator = CreateVerticalOrHorizontalRangeNode;

            // File is only used in Reference and not directly, so don't use NotSupportedNode since it is never evaluated.
            grammar.File.AstConfig.NodeCreator = DontCreateNode;
            grammar.File.SetFlag(TermFlags.AstDelayChildren);

            grammar.UDFunctionCall.AstConfig.NodeCreator = CreateNotSupportedNode("custom functions");
            grammar.UDFName.AstConfig.NodeCreator = DontCreateNode;
            grammar.UDFToken.AstConfig.NodeCreator = DontCreateNode;

            grammar.StructuredReference.AstConfig.NodeCreator = CreateNotSupportedNode("structured references");
            grammar.StructuredReference.SetFlag(TermFlags.AstDelayChildren);

            // Irony has a few bugs. If it throws a NRE in BuildAst(parseNode), some node is missing a setting to create node for the term.
            grammar.LanguageFlags |= LanguageFlags.CreateAst;
            return grammar;
        }

        private void DontCreateNode(AstContext context, ParseTreeNode parseNode)
        {
            // Don't create an AST node for the parseNode. Its children will use their AstConfig to create their AST nodes.
        }

        private void CreateNumberNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new Expression(parseNode.Token.Value);
        }

        private void CreateBoolNode(AstContext context, ParseTreeNode parseNode)
        {
            var valueString = parseNode.Token.ValueString;
            var boolValue = string.Equals(valueString, "TRUE", StringComparison.OrdinalIgnoreCase);
            parseNode.AstNode = new Expression(boolValue);
        }

        private void CreateTextNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = new Expression(parseNode.Token.ValueString);
        }

        private void CreateErrorNode(AstContext context, ParseTreeNode parseNode)
        {
            var errorType = ErrorMap[parseNode.ChildNodes.Single().Token.ValueString];
            parseNode.AstNode = new ErrorExpression(errorType);
        }

        private void CreateReferenceNode(AstContext context, ParseTreeNode parseNode)
        {
            if (HasMatchingChildren(parseNode, GrammarNames.UDFunctionCall))
            {
                parseNode.AstNode = parseNode.ChildNodes.Single().AstNode;
                return;
            }

            if (parseNode.ChildNodes.Count == 1)
            {
                var firstNode = parseNode.ChildNodes[0];
                if (ReferenceItemTerms.Contains(firstNode.Term.Name))
                {
                    parseNode.AstNode = firstNode.AstNode;
                    return;
                }
                else if (firstNode.Term.Name == GrammarNames.ReferenceFunctionCall)
                {
                    parseNode.AstNode = firstNode.AstNode;
                    return;
                }
                else if (firstNode.Term.Name == GrammarNames.Reference)
                {
                    // another reference in parenthesis
                    parseNode.AstNode = firstNode.AstNode;
                    return;
                }
            }
            else if (parseNode.ChildNodes.Count == 2
                && parseNode.ChildNodes[0].Term.Name == GrammarNames.Prefix
                && ReferenceItemTerms.Contains(parseNode.ChildNodes[1].Term.Name))
            {
                // prefix nebo expression
                var prefixResult = GetPrefix(parseNode.ChildNodes[0]);
                if (prefixResult.Item2 is not null)
                {
                    parseNode.AstNode = prefixResult.Item2;
                    return;
                }

                var addressResult = GetReferenceItemAddress(parseNode.ChildNodes[1]);
                if (addressResult.Item2 is not null)
                {
                    parseNode.AstNode = prefixResult.Item2;
                    return;
                }

                parseNode.AstNode = CreateExternalExpression(prefixResult.Item1 + addressResult.Item1);
                return;
            }

            throw new NotImplementedException();

            static Tuple<string, NotSupportedNode> GetReferenceItemAddress(ParseTreeNode referenceItemUnion)
            {
                if (referenceItemUnion.Term.Name == GrammarNames.Cell)
                {
                    return new Tuple<string, NotSupportedNode>(referenceItemUnion.ChildNodes[0].Token.ValueString, null);
                }

                throw new NotImplementedException();
            }
        }

        private static Tuple<string, NotSupportedNode> GetPrefix(ParseTreeNode prefixNode)
        {
            if (HasMatchingChildren(prefixNode, GrammarNames.TokenSheet))
            {
                return new Tuple<string, NotSupportedNode>(prefixNode.ChildNodes.Single().Token.ValueString, null);
            }

            if (HasMatchingChildren(prefixNode, "'", GrammarNames.TokenSheetQuoted))
            {
                return new Tuple<string, NotSupportedNode>("'" + prefixNode.ChildNodes[1].Token.ValueString, null);
            }

            if (HasMatchingChildren(prefixNode, GrammarNames.File, GrammarNames.TokenSheet))
            {
                return new Tuple<string, NotSupportedNode>(null, new NotSupportedNode("external reference"));
            }

            throw new NotImplementedException();
        }

        private void CreateCellNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = CreateExternalExpression(parseNode.Token.ValueString);
        }

        private void CreateFunctionCallNode(AstContext context, ParseTreeNode parseNode)
        {
            if (parseNode.ChildNodes.Count == 2)
            {
                var firstTermName = parseNode.ChildNodes[0].Term.Name;
                var secondTermName = parseNode.ChildNodes[1].Term.Name;
                if ((firstTermName == "-" || firstTermName == "+" || firstTermName == "@") && secondTermName == GrammarNames.Formula)
                {
                    parseNode.AstNode = new UnaryExpression(firstTermName, (Expression)parseNode.ChildNodes[1].AstNode);
                    return;
                }
                else if (firstTermName == GrammarNames.FunctionName
                    && secondTermName == GrammarNames.Arguments)
                {
                    parseNode.AstNode = CreateExcelFunctionCallExpression(parseNode.ChildNodes[0], parseNode.ChildNodes[1]);
                    return;
                }
            }
            else if (parseNode.ChildNodes.Count == 3)
            {
                var middleTerm = parseNode.ChildNodes[1].Term;

                if (_binaryOpMap.TryGetValue(middleTerm, out var infixOp)
                    && parseNode.ChildNodes[0].Term.Name == GrammarNames.Formula
                    && parseNode.ChildNodes[2].Term.Name == GrammarNames.Formula)
                {
                    parseNode.AstNode = new BinaryExpression(infixOp, (Expression)parseNode.ChildNodes[0].AstNode, (Expression)parseNode.ChildNodes[2].AstNode);
                    return;
                }
            }

            throw new NotSupportedException();
        }

        private void CreateReferenceFunctionCallNode(AstContext context, ParseTreeNode parseNode)
        {
            // Has to be first to have higher priority than reference range operator
            if (IsLegacyRange(parseNode, out var rangeExpression))
            {
                parseNode.AstNode = rangeExpression;
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Reference, ":", GrammarNames.Reference))
            {
                parseNode.AstNode = new NotSupportedNode("binary range operator");
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Reference, GrammarNames.TokenIntersect, GrammarNames.Reference))
            {
                parseNode.AstNode = new NotSupportedNode("range intersection operator");
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Union))
            {
                parseNode.AstNode = parseNode.ChildNodes.Single().AstNode;
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.RefFunctionName, GrammarNames.Arguments))
            {
                parseNode.AstNode = CreateExcelFunctionCallExpression(parseNode.ChildNodes[0], parseNode.ChildNodes[1]);
                return;
            }

            if (HasMatchingChildren(parseNode, GrammarNames.Reference, "#"))
            {
                parseNode.AstNode = new NotSupportedNode("spill range operator");
                return;
            }

            throw new NotSupportedException();
        }

        private Expression CreateExcelFunctionCallExpression(ParseTreeNode nameNode, ParseTreeNode argumentsNode)
        {

            var nameWithOpeningBracket = nameNode.ChildNodes.Single().Token.ValueString;
            var functionName = nameWithOpeningBracket.Substring(0, nameWithOpeningBracket.Length - 1);
            var foundFunction = _fnTbl.TryGetValue(functionName, out FunctionDefinition functionDefinition);
            if (!foundFunction && functionName.StartsWith($"{defaultFunctionNameSpace}."))
                foundFunction = _fnTbl.TryGetValue(functionName.Substring(defaultFunctionNameSpace.Length + 1), out functionDefinition);

            if (!foundFunction)
                throw new NameNotRecognizedException($"The function `{functionName}` was not recognised.");

            var arguments = argumentsNode.ChildNodes.Select(treeNode => treeNode.AstNode).Cast<Expression>().ToList();
            return new FunctionExpression(functionDefinition, arguments);
        }

        private void CreateNamedRangeNode(AstContext context, ParseTreeNode parseNode)
        {
            // Named range can be NameToken or NamedRangeCombinationToken. The second one is there only to detect names like A1A1.
            var rangeName = parseNode.ChildNodes.Single().Token.ValueString;
            parseNode.AstNode = CreateExternalExpression(rangeName);
        }

        private void CreateVerticalOrHorizontalRangeNode(AstContext context, ParseTreeNode parseNode)
        {
            parseNode.AstNode = CreateExternalExpression(parseNode.Token.ValueString);
        }

        private static AstNodeCreator CreateCopyNode(int childIndex)
        {
            return (context, parseNode) =>
            {
                var copyNode = parseNode.ChildNodes[childIndex];
                parseNode.AstNode = copyNode.AstNode;
            };
        }

        private static AstNodeCreator CreateNotSupportedNode(string featureText)
        {
            return (_, parseNode) => parseNode.AstNode = new NotSupportedNode(featureText);
        }

        #region Old parser compatibility methods

        /// <summary>
        /// Old parser didn't have any range operations and was only able to parse certain patterns of range operations. This is here to keep
        /// it working, until we get range operations working.
        /// </summary>
        private bool IsLegacyRange(ParseTreeNode referenceFunctionCall, out Expression rangeExpression)
        {
            if (HasMatchingChildren(referenceFunctionCall, GrammarNames.Reference, ":", GrammarNames.Reference))
            {
                var leftReference = referenceFunctionCall.ChildNodes[0];
                var rightReference = referenceFunctionCall.ChildNodes[2];
                if (HasMatchingChildren(leftReference, GrammarNames.Cell) && HasMatchingChildren(rightReference, GrammarNames.Cell))
                {
                    // Pattern A1:B1
                    var range = leftReference.ChildNodes.Single().ChildNodes.Single().Token.ValueString
                        + ":" + rightReference.ChildNodes.Single().ChildNodes.Single().Token.ValueString;
                    rangeExpression = CreateExternalExpression(range);
                    return true;
                }

                if (HasMatchingChildren(leftReference, GrammarNames.Prefix, GrammarNames.Cell) && HasMatchingChildren(rightReference, GrammarNames.Cell))
                {
                    // Pattern Sheet1!A1:B1
                    var prefixNode = leftReference;
                    string sheet;
                    if (HasMatchingChildren(prefixNode, GrammarNames.TokenSheet))
                    {
                        sheet = prefixNode.ChildNodes.Single().Token.ValueString;
                    }
                    else if (HasMatchingChildren(prefixNode, GrammarNames.TokenSheetQuoted))
                    {
                        sheet = prefixNode.ChildNodes.Single().Token.ValueString;
                    }
                    else
                    {
                        rangeExpression = new NotSupportedNode("file reference");
                        return true;
                    }

                    var range = sheet
                        + "!" + leftReference.ChildNodes[1].ChildNodes.Single().Token.ValueString
                        + ":" + rightReference.ChildNodes.Single().ChildNodes.Single().Token.ValueString;
                    rangeExpression = CreateExternalExpression(range);
                    return true;
                }
            }

            rangeExpression = null;
            return false;
        }

        private XObjectExpression CreateExternalExpression(string referenceOrNamedRange)
        {
            // TODO: This is a wrong way to create AST, because it doesn't separate parsing and evaluation, but throws exceptions during parsing.
            // Example: =NonExistentSheet!A1 is a valid formula that should return #REF! or if a 'NonExistentSheet' is later added, it should return the value of a cell.
            // Kept for compatibility with the old parser.
            var xObj = _engine.GetExternalObject(referenceOrNamedRange);
            if (xObj == null)
                throw new NameNotRecognizedException($"The identifier `{referenceOrNamedRange}` was not recognised.");

            return new XObjectExpression(xObj);
        }

        private void CreateEmptyArgumentNode(AstContext context, ParseTreeNode parseNode)
        {
            // TODO: This is useless for AST, but kept for compatibility reasons with old parser and some function that use it.
            parseNode.AstNode = new EmptyValueExpression();
        }

        #endregion

        private static bool HasMatchingChildren(ParseTreeNode node, params string[] termNames)
        {
            return node.ChildNodes.Select(c => c.Term.Name).SequenceEqual(termNames);
        }
    }
}
