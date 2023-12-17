using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public static class MSAccessAnalyzer
    {
        public static int CollectCommands(
            this RuleContext context,
            CaretPosition caretPosition,
            string tokenSeparator,
            int commandSeparatorTokenType,
            out IList<ParsedTreeCommand> commands)
        {
            commands = new List<ParsedTreeCommand>();
            commands.Add(new ParsedTreeCommand
            {
                StartOffset = -1,
                StopOffset = -1,
                StartLine = -1,
                StopLine = -1,
                StartColumn = -1,
                StopColumn = -1,
                Text = "",
                Context = null
            });
            return _CollectCommands(context, caretPosition, tokenSeparator, commandSeparatorTokenType, commands, -1, null);
        }

        private static IToken GetAsSymbolOfType(this IParseTree context, int symbolType)
        {
            while (context is IRuleNode ruleNode && ruleNode.ChildCount == 1)
            {
                context = ruleNode.GetChild(0);
            }
            if (context is ITerminalNode terminalNode && terminalNode.Symbol.Type == symbolType)
            {
                return terminalNode.Symbol;
            }
            return null;
        }

        private static bool HasAndOrExprWithoutParens(IParseTree context)
        {
            if (context is MSAccessParser.Select_stmtContext)
                return false;
            if (context is MSAccessParser.ExprContext thisCtx && context.Parent is MSAccessParser.ExprContext parentCtx &&
                thisCtx.op != null && parentCtx.op != null &&
                thisCtx.op.Type.In(MSAccessParser.AND_, MSAccessParser.OR_) &&
                parentCtx.op.Type.In(MSAccessParser.AND_, MSAccessParser.OR_) &&
                thisCtx.op.Type != parentCtx.op.Type)
            {
                return true;
            }
            for (var n = 0; n < context.ChildCount; ++n)
            {
                var child = context.GetChild(n);
                var result = HasAndOrExprWithoutParens(child);
                if (result) return true;
            }
            return false;
        }

        private static bool HasAggregateFunction(IParseTree context)
        {
            if (context is MSAccessParser.Select_stmtContext)
                return false;
            if (context is MSAccessParser.Function_exprContext ctx &&
                ctx.functionName.GetText().ToLower().In("sum", "avg", "min", "max", "count"))
            {
                return true;
            }

            for (var n = 0; n < context.ChildCount; ++n)
            {
                var child = context.GetChild(n);
                var result = HasAggregateFunction(child);
                if (result) return true;
            }
            return false;
        }

        private static bool HasMultipleWheres(IParseTree context)
        {
            if (string.Equals(context.GetText(), "where", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            } 
            for (var n = 0; n < context.ChildCount; ++n)
            {
                var child = context.GetChild(n);
                var result = HasMultipleWheres(child);
                if (result) return true;
            }
            return false;
        }

        private static bool HasWhereClause(IParseTree context)
        {
            if (context is MSAccessParser.Where_clauseContext)
            {
                return true;
            } 
            for (var n = 0; n < context.ChildCount; ++n)
            {
                var child = context.GetChild(n);
                var result = HasWhereClause(child);
                if (result) return true;
            }
            return false;
        }
        public static IParseTree FindParentOfAnyType(IParseTree context, IList<Type> targets)
        {
            var parent = context.Parent;
            if (parent == null)
            {
                return null;
            }
            foreach (Type target in targets)
            {
                if (target.IsAssignableFrom(parent.GetType()))
                {
                    return parent;
                }
            }
            var result = FindParentOfAnyType(parent, targets);
            if (result != null)
            {
                return result;
            }
            return null;
        }

        private static void FindTopWarnings(MSAccessParser.Select_core_stmtContext ctx, ParsedTreeCommand command)
        {
            if (ctx.orderByClause == null)
            {
                command.AddWarning(ctx, ParserMessageType.TOP_KEYWORD_WITHOUT_ORDER_BY_CLAUSE);
                //IParseTree exprParent = FindParentOfAnyType(ctx, new List<Type> { typeof(MSAccessParser.ExprContext) });
                //IParseTree insertStmtParent = FindParentOfAnyType(ctx, new List<Type> { typeof(MSAccessParser.Insert_stmtContext) });
                //if ((exprParent != null && exprParent is MSAccessParser.ExprContext expr && expr.subquery == ctx.Parent) || 
                //    (insertStmtParent != null && insertStmtParent is MSAccessParser.Insert_stmtContext insrt && insrt.subquery == ctx))
                //{
                //}
            }
            if (ctx.selectClause?.limit != null)
            {
                try
                {
                    int result = Int32.Parse(ctx.selectClause?.limit.Text);
                    if (result == 1 && ctx.selectClause?.percent == null)
                    {
                        IParseTree exprParent = FindParentOfAnyType(ctx, new List<Type> { typeof(MSAccessParser.ExprContext) });
                        IParseTree insertStmtParent = FindParentOfAnyType(ctx, new List<Type> { typeof(MSAccessParser.Insert_stmtContext) });
                        if ((exprParent != null && exprParent is MSAccessParser.ExprContext expr && expr.subquery == ctx.Parent) ||
                            (insertStmtParent != null && insertStmtParent is MSAccessParser.Insert_stmtContext insrt && insrt.subquery == ctx))
                        {
                            IParseTree parent = FindParentOfAnyType(ctx, new List<Type> { typeof(MSAccessParser.Where_clauseContext) });
                            if (parent != null && parent is MSAccessParser.Where_clauseContext where_ClauseContext)
                            {
                                if (where_ClauseContext.whereExpr == null ||
                                    (where_ClauseContext.whereExpr != null &&
                                    where_ClauseContext.whereExpr.selector == null && where_ClauseContext.whereExpr?.op?.Type != MSAccessParser.IN_))
                                {
                                    command.AddWarning(ctx, ParserMessageType.TOP_KEYWORD_MIGHT_RETURN_MULTIPLE_ROWS);
                                }
                            }
                            else
                            {
                                command.AddWarning(ctx, ParserMessageType.TOP_KEYWORD_MIGHT_RETURN_MULTIPLE_ROWS);
                            }
                        }
                    }
                    if (result < 1 && ctx.selectClause?.percent == null)
                    {
                        command.AddWarning(ctx, ParserMessageType.TOP_LIMIT_CONSTRAINT);
                    }
                    if ((result < 1 || result > 100) && ctx.selectClause?.percent != null)
                    {
                        command.AddWarning(ctx, ParserMessageType.TOP_LIMIT_PERCENT_CONSTRAINT);
                    }
                }
                catch (FormatException)
                {
                    command.AddWarning(ctx, ParserMessageType.POSSIBLE_NON_INTEGER_VALUE_WITH_TOP);
                }
            }
            IList<MSAccessParser.Result_columnContext> _resultColumns = ctx?.selectClause._resultColumns;
            if (_resultColumns.Count == 1) 
            {
                foreach (MSAccessParser.Result_columnContext column in _resultColumns)
                {

                    if (column?.columnExpr?.functionExpr?.functionName?.GetText() != null
                        && column.columnExpr.functionExpr.functionName.GetText().ToLower().In("sum", "avg", "min", "max", "count"))
                    {
                        command.AddWarning(ctx, ParserMessageType.ONE_ROW_IN_RESULT_WITH_TOP);
                        continue;
                    }
                }
            }
        }

        private static bool IsLogicalExpression(IParseTree context)
        {
            if (!(context is MSAccessParser.ExprContext ctx)) return false;

            if (ctx.op == null)
            {
                return ctx.subquery == null && (
                    ctx.literalExpr == null || ctx.literalExpr.literal.Type.In(MSAccessParser.FALSE_, MSAccessParser.TRUE_));
            }

            return !ctx.op.Type.In(
                MSAccessParser.STAR, MSAccessParser.DIV, MSAccessParser.IDIV, MSAccessParser.MOD_,
                MSAccessParser.AMP, MSAccessParser.PLUS, MSAccessParser.MINUS
            );
        }

        private static bool In<T>(this T x, params T[] set)
        {
            return set.Contains(x);
        }

        private static void _AnalyzeToken(IToken token, ParsedTreeCommand command)
        {
            switch (token.Type)
            {
                case MSAccessParser.IDENTIFIER:
                    {
                        if (token.Text[0] == '"' && token.Text[token.Text.Length - 1] == '"')
                            command.AddWarning(token, ParserMessageType.DOUBLE_QUOTES);
                        break;
                    }
            }
        }

        private static void _AnalyzeRuleContext(RuleContext context, ParsedTreeCommand command)
        {
            switch (context.RuleIndex)
            {
                case MSAccessParser.RULE_select_into_stmt:
                    {
                        if (context is MSAccessParser.Select_into_stmtContext ctx)
                        {
                            if (ctx.selectClause?.distinct != null && ctx.groupByClause != null)
                                command.AddWarning(ctx, ParserMessageType.DISTINCT_KEYWORD_WITH_GROUP_BY_CLAUSE);

                            var tables = ctx.fromClause?._tables;
                            var joins = ctx._joinClause;
                            var columns = ctx.selectClause?._resultColumns;
                            if ((tables != null && tables.Count > 1 || joins != null && joins.Count > 0) && columns != null &&
                                columns.Any(c => c.STAR() != null))
                                command.AddWarning(ctx, ParserMessageType.SELECT_ALL_WITH_MULTIPLE_JOINS);
                        }
                        break;
                    }
                case MSAccessParser.RULE_select_stmt:
                    {
                        if (context is MSAccessParser.Select_stmtContext ctx && ctx._statements.Count > 1)
                        {
                            foreach (var statement in ctx._statements)
                            {
                                var columns = statement.selectClause?._resultColumns;
                                if (columns != null && columns.Any(c => c.STAR() != null || c.prefixed_star() != null))
                                    command.AddWarning(statement, ParserMessageType.SELECT_ALL_IN_UNION_STATEMENT);
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_select_core_stmt:
                    {
                        if (context is MSAccessParser.Select_core_stmtContext ctx)
                        {
                            if (ctx.selectClause.distinct != null && ctx.groupByClause != null)
                            {
                                command.AddWarning(ctx, ParserMessageType.DISTINCT_KEYWORD_WITH_GROUP_BY_CLAUSE);
                            }
                            if (ctx.selectClause.top != null)
                            {
                                FindTopWarnings(ctx, command);
                            }
                            if (ctx.groupByClause == null &&
                                ctx.selectClause._resultColumns.Any(c => c.columnExpr?.prefixedColumnName != null) &&
                                HasAggregateFunction(ctx.selectClause))
                            {
                                command.AddWarning(ctx.selectClause, ParserMessageType.AGGREGATE_FUNCTION_WITHOUT_GROUP_BY_CLAUSE);
                            }

                            var tables = ctx.fromClause?._tables;
                            var joins = ctx._joinClause;
                            var columns = ctx.selectClause._resultColumns;
                            if ((tables != null && tables.Count > 1 || joins != null && joins.Count > 0) && columns != null &&
                                columns.Any(c => c.STAR() != null))
                            {
                                command.AddWarning(ctx, ParserMessageType.SELECT_ALL_WITH_MULTIPLE_JOINS);
                            }

                            if (joins != null && joins.Count > 0 && joins.Any(j => j.LEFT_() != null || j.RIGHT_() != null)
                                && columns != null && columns.Any(c => c.columnExpr?.functionExpr is var f && f != null &&
                                                                       f.functionName.GetText().ToLower() == "count" &&
                                                                       f.STAR() != null))
                            {
                                command.AddWarning(ctx, ParserMessageType.COUNT_FUNCTION_WITH_OUTER_JOIN);
                            }
                            if (ctx.groupByClause != null && columns != null 
                                && columns.Count - 1 != ctx.groupByClause._groupingTerms.Count 
                                && columns.Count != ctx.groupByClause._groupingTerms.Count)
                            {
                                command.AddWarning(ctx, ParserMessageType.MISSING_COLUMN_IN_GROUP_BY_CLAUSE);
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_insert_stmt:
                    {
                        if (context is MSAccessParser.Insert_stmtContext ctx)
                        {
                            if (ctx._columnNames.Count == 0)
                                command.AddWarning(ctx, ParserMessageType.INSERT_STATEMENT_WITHOUT_COLUMN_NAMES);

                            var columns = ctx.subquery?.selectClause?._resultColumns;
                            if (columns != null && columns.Any(c => c.STAR() != null || c.prefixed_star() != null))
                                command.AddWarning(ctx.subquery, ParserMessageType.SELECT_ALL_IN_INSERT_STATEMENT);
                        }
                        break;
                    }
                case MSAccessParser.RULE_order_by_clause:
                    {
                        if (context is MSAccessParser.Order_by_clauseContext ctx)
                        {
                            foreach (var orderingTerm in ctx._orderingTerms)
                            {
                                if (orderingTerm.orderingExpr.literalExpr == null) continue;

                                command.AddWarning(ctx, ParserMessageType.ORDERING_BY_ORDINAL);
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_group_by_clause:
                    {
                        if (context is MSAccessParser.Group_by_clauseContext ctx)
                        {
                            if (ctx.havingExpr != null && !HasAggregateFunction(ctx.havingExpr))
                                command.AddWarning(ctx.havingExpr, ParserMessageType.HAVING_CLAUSE_WITHOUT_AGGREGATE_FUNCTION);
                            if (ctx.havingExpr != null && HasAndOrExprWithoutParens(ctx.havingExpr))
                                command.AddWarning(ctx.havingExpr, ParserMessageType.AND_OR_MISSING_PARENTHESES_IN_WHERE_CLAUSE);
                            if (ctx.havingExpr != null && !IsLogicalExpression(ctx.havingExpr))
                                command.AddWarning(ctx.havingExpr, ParserMessageType.NOT_LOGICAL_OPERAND);

                            foreach (var groupingTerm in ctx._groupingTerms)
                            {
                                if (!HasAggregateFunction(groupingTerm)) continue;

                                command.AddWarning(ctx, ParserMessageType.AGGREGATE_FUNCTION_IN_GROUP_BY_CLAUSE);
                                break;
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_select_clause:
                    {
                        if (context is MSAccessParser.Select_clauseContext ctx)
                        {
                            for (var i = 0; i < ctx._resultColumns.Count; ++i)
                            {
                                var firstName = ctx._resultColumns[i].columnExpr?.prefixedColumnName?.columnName;
                                var firstAlias = ctx._resultColumns[i].columnAlias;
                                var firstText = firstAlias?.GetText() ?? firstName?.GetText();
                                if (firstText == null)
                                    continue;

                                for (var j = i + 1; j < ctx._resultColumns.Count; ++j)
                                {
                                    var secondName = ctx._resultColumns[j].columnExpr?.prefixedColumnName?.columnName;
                                    var secondAlias = ctx._resultColumns[j].columnAlias;
                                    var secondText = secondAlias?.GetText() ?? secondName?.GetText();
                                    if (secondText == null || firstText != secondText)
                                        continue;

                                    command.AddWarning(ctx, ParserMessageType.DUPLICATE_SELECTED_COLUMN_IN_SELECT_CLAUSE);
                                }
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_where_clause:
                    {
                        if (context is MSAccessParser.Where_clauseContext ctx)
                        {
                            if (HasAndOrExprWithoutParens(ctx.whereExpr))
                                command.AddWarning(ctx, ParserMessageType.AND_OR_MISSING_PARENTHESES_IN_WHERE_CLAUSE);
                            if (HasAggregateFunction(ctx))
                                command.AddWarning(ctx, ParserMessageType.AGGREGATE_FUNCTION_IN_WHERE_CLAUSE);
                            if (!IsLogicalExpression(ctx.whereExpr))
                                command.AddWarning(ctx, ParserMessageType.NOT_LOGICAL_OPERAND);
                            if (HasMultipleWheres(ctx.whereExpr))
                                command.AddWarning(ctx, ParserMessageType.MULTIPLE_WHERE_USED);
                        }
                        break;
                    }
                case MSAccessParser.RULE_result_column:
                    {
                        if (context is MSAccessParser.Result_columnContext ctx)
                        {
                            // there's more difficult case, where outer select statement has asterisk,
                            // but inner statement has expression without alias
                            if (ctx.Parent is MSAccessParser.Select_clauseContext &&
                                (ctx.Parent.Parent is MSAccessParser.Select_into_stmtContext &&
                                 ctx.Parent.Parent.Parent is MSAccessParser.Sql_stmtContext ||
                                 ctx.Parent.Parent is MSAccessParser.Select_core_stmtContext &&
                                 ctx.Parent.Parent.Parent is MSAccessParser.Select_stmtContext &&
                                 ctx.Parent.Parent.Parent.Parent is MSAccessParser.Sql_stmtContext) &&
                                ctx.columnExpr != null && ctx.columnAlias == null && ctx.columnExpr.prefixedColumnName == null)
                                command.AddWarning(ctx, ParserMessageType.MISSING_COLUMN_ALIAS_IN_SELECT_CLAUSE);
                        }
                        break;
                    }
                case MSAccessParser.RULE_function_expr:
                    {
                        if (context is MSAccessParser.Function_exprContext ctx)
                        {
                            if (ctx.functionName.GetText().ToLower() == "sum" && ctx._params.Count == 1 && ctx._params[0].GetText() == "1")
                                command.AddWarning(ctx, ParserMessageType.USE_COUNT_FUNCTION);
                        }
                        break;
                    }
                case MSAccessParser.RULE_join_clause:
                    {
                        if (context is MSAccessParser.Join_clauseContext ctx)
                        {
                            if (ctx.on == null && ctx.expression == null)
                            {
                                command.AddWarning(ctx, ParserMessageType.MISSING_EXPRESSION_IN_JOIN_CLAUSE);
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_from_clause:
                    {
                        if (context is MSAccessParser.From_clauseContext ctx)
                        {
                            if (ctx?._tables.Count > 1 && !HasWhereClause(ctx))
                            {
                                command.AddWarning(ctx, ParserMessageType.MISSING_EXPRESSION_IN_JOIN_CLAUSE);
                            }
                            if (string.Equals(ctx?.Stop?.Text, "where", StringComparison.OrdinalIgnoreCase))
                            {
                                command.AddWarning(ctx, ParserMessageType.MISSING_EXPRESSION_IN_WHERE_CLAUSE);
                            }
                            if (ctx?._table_or_subquery != null &&
                                ctx._table_or_subquery.Start.Type == MSAccessParser.OPEN_PAR &&
                                ctx._table_or_subquery.Stop.Type == MSAccessParser.CLOSE_PAR &&
                                ctx._table_or_subquery.table_alias() == null)
                            {
                                command.AddWarning(ctx, ParserMessageType.MISSING_ALIAS_IN_FROM_SUBQUERY);
                            }
                        }
                        break;
                    }
                case MSAccessParser.RULE_expr:
                    {
                        if (context is MSAccessParser.ExprContext ctx && ctx.op != null)
                        {
                            if (ctx.subquery != null)
                            {
                                foreach (var statement in ctx.subquery._statements)
                                {
                                    var columns = statement.selectClause?._resultColumns;
                                    if (statement.selectClause?.limit == null && statement.orderByClause != null)
                                        command.AddWarning(ctx.subquery, ParserMessageType.ORDER_BY_CLAUSE_IN_SUB_QUERY_WITHOUT_LIMIT);
                                    if ((ctx.op.Type == MSAccessParser.IN_ || ctx.selector != null) &&
                                        columns != null && columns.Any(c => c.STAR() != null || c.prefixed_star() != null))
                                        command.AddWarning(ctx.subquery, ParserMessageType.SELECT_ALL_IN_SUB_QUERY);
                                    if (ctx.op.Type != MSAccessParser.EXISTS_ && columns != null && columns.Count > 1)
                                        command.AddWarning(ctx.subquery, ParserMessageType.MULTIPLE_COLUMNS_IN_SUB_QUERY);
                                }
                            }
                            switch (ctx.op.Type)
                            {
                                case MSAccessParser.ALIKE_:
                                case MSAccessParser.LIKE_:
                                    {
                                        var rhs = ctx.rhs.GetAsSymbolOfType(MSAccessParser.STRING_LITERAL);
                                        if (rhs != null && !rhs.Text.Contains("%") && !rhs.Text.Contains("_"))
                                            command.AddWarning(ctx, ParserMessageType.MISSING_WILDCARDS_IN_LIKE_EXPRESSION);
                                    }
                                    {
                                        var lhs = ctx.lhs.prefixedColumnName;
                                        var rhs = ctx.rhs.prefixedColumnName;
                                        if (lhs != null && rhs != null)
                                            command.AddWarning(ctx, ParserMessageType.COLUMN_LIKE_COLUMN);
                                        break;
                                    }
                                case MSAccessParser.LT:
                                case MSAccessParser.LT_EQ:
                                case MSAccessParser.GT:
                                case MSAccessParser.GT_EQ:
                                case MSAccessParser.EQ:
                                case MSAccessParser.NOT_EQ1:
                                case MSAccessParser.NOT_EQ2:
                                    {
                                        var lhs = ctx.lhs.GetAsSymbolOfType(MSAccessParser.NULL_);
                                        var rhs = ctx.rhs.GetAsSymbolOfType(MSAccessParser.NULL_);
                                        if (lhs != null || rhs != null)
                                            command.AddWarning(ctx, ParserMessageType.EQUALITY_WITH_NULL);

                                        if (ctx.op.Type == MSAccessParser.EQ &&
                                            ctx.selector != null && ctx.selector.Type == MSAccessParser.ALL_)
                                            command.AddWarning(ctx, ParserMessageType.EQUALS_ALL);

                                        if (ctx.op.Type == MSAccessParser.NOT_EQ2 &&
                                            ctx.selector != null && ctx.selector.Type.In(MSAccessParser.ANY_, MSAccessParser.SOME_))
                                            command.AddWarning(ctx, ParserMessageType.NOT_EQUALS_ANY);

                                        if (!ctx.op.Type.In(MSAccessParser.EQ, MSAccessParser.NOT_EQ1, MSAccessParser.NOT_EQ2))
                                            break;

                                        lhs = ctx.lhs.GetAsSymbolOfType(MSAccessParser.STRING_LITERAL);
                                        rhs = ctx.rhs.GetAsSymbolOfType(MSAccessParser.STRING_LITERAL);
                                        if (lhs != null && (lhs.Text.Contains("%") || lhs.Text.Contains("_")) ||
                                            rhs != null && (rhs.Text.Contains("%") || rhs.Text.Contains("_")))
                                            command.AddWarning(ctx, ParserMessageType.EQUALITY_WITH_TEXT_PATTERN);
                                        break;
                                    }
                                case MSAccessParser.XOR_:
                                case MSAccessParser.AND_:
                                case MSAccessParser.OR_:
                                case MSAccessParser.EQV_:
                                    {
                                        if (!IsLogicalExpression(ctx.lhs))
                                            command.AddWarning(ctx.lhs, ParserMessageType.NOT_LOGICAL_OPERAND);
                                        if (!IsLogicalExpression(ctx.rhs))
                                            command.AddWarning(ctx.rhs, ParserMessageType.NOT_LOGICAL_OPERAND);
                                        break;
                                    }
                                case MSAccessParser.DIV:
                                    {
                                        var lhs = ctx.lhs.functionExpr;
                                        var rhs = ctx.rhs.functionExpr;
                                        if (lhs != null && rhs != null &&
                                            lhs.functionName.GetText().ToLower() == "sum" &&
                                            rhs.functionName.GetText().ToLower() == "count")
                                            command.AddWarning(ctx, ParserMessageType.USE_AVG_FUNCTION);
                                        break;
                                    }
                            }
                        }
                        break;
                    }
            }
        }

        private static int _CollectCommands(
            RuleContext context,
            CaretPosition caretPosition,
            string tokenSeparator,
            int commandSeparatorTokenType,
            in IList<ParsedTreeCommand> commands,
            int enclosingCommandIndex,
            in IList<StringBuilder> functionParams)
        {

            if (context is MSAccessParser.Sql_stmtContext && commands.Last().Context == null)
            {
                commands.Last().Context = context; // base starting branch
            }
            _AnalyzeRuleContext(context, commands.Last());
            for (var i = 0; i < context.ChildCount; ++i)
            {
                var child = context.GetChild(i);
                if (child is ITerminalNode terminalNode)
                {
                    var token = terminalNode.Symbol;
                    _AnalyzeToken(token, commands.Last());
                    var tokenLength = token.StopIndex - token.StartIndex + 1;
                    if (token.Type == TokenConstants.EOF || token.Type == commandSeparatorTokenType)
                    {
                        if (enclosingCommandIndex == -1 && (
                            caretPosition.Line > commands.Last().StartLine && caretPosition.Line < token.Line ||
                            caretPosition.Line == commands.Last().StartLine && caretPosition.Column >= commands.Last().StartColumn && (caretPosition.Line < token.Line || caretPosition.Column <= token.Column) ||
                            caretPosition.Line == token.Line && caretPosition.Column <= token.Column))
                        {
                            enclosingCommandIndex = commands.Count - 1;
                        }

                        if (token.Type == TokenConstants.EOF)
                        {
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(commands.Last().Text))
                        {
                            commands.Last().AddWarning(token, ParserMessageType.UNNECESSARY_SEMICOLON);
                        }
                        commands.Add(new ParsedTreeCommand());
                    }
                    else
                    {
                        if (commands.Last().StartOffset == -1)
                        {
                            commands.Last().StartLine = token.Line;
                            commands.Last().StartColumn = token.Column;
                            commands.Last().StartOffset = token.StartIndex;
                        }
                        commands.Last().StopLine = token.Line;
                        commands.Last().StopColumn = token.Column + tokenLength;
                        commands.Last().StopOffset = token.StopIndex;

                        if (functionParams is null)
                        {
                            commands.Last().Text = commands.Last().Text + child.GetText() + tokenSeparator;
                        }
                        else if (context.RuleIndex != MSAccessParser.RULE_function_expr || token.Type != MSAccessParser.OPEN_PAR && token.Type != MSAccessParser.CLOSE_PAR && token.Type != MSAccessParser.COMMA)
                        {
                            functionParams.Last().Append(child.GetText() + tokenSeparator);
                        }
                        if (token.Type == MSAccessParser.TOP_ && i == context.ChildCount - 1) 
                        {
                            commands.Last().AddWarning(token, ParserMessageType.POSSIBLE_NON_INTEGER_VALUE_WITH_TOP);
                        }
                    }
                }
                else
                {
                    var ctx = child as RuleContext;
                    if (ctx?.RuleIndex == MSAccessParser.RULE_optional_parens) continue;

                    if (ctx?.RuleIndex == MSAccessParser.RULE_prefixed_star ||
                        ctx?.RuleIndex == MSAccessParser.RULE_prefixed_column_name)
                    {
                        enclosingCommandIndex = _CollectCommands(ctx, caretPosition, "", commandSeparatorTokenType, commands, enclosingCommandIndex, functionParams);
                        if (functionParams is null)
                            commands.Last().Text += tokenSeparator;
                        else
                            functionParams.Last().Append(tokenSeparator);
                    }
                    else if (ctx?.RuleIndex == MSAccessParser.RULE_function_expr)
                    {
                        var p = new List<StringBuilder> { new StringBuilder() };
                        enclosingCommandIndex = _CollectCommands(ctx, caretPosition, tokenSeparator, commandSeparatorTokenType, commands, enclosingCommandIndex, p);
                        var functionName = p[0].ToString().ToLower();
                        p.RemoveAt(0);
                        var functionCallString = "";
                        if (functionName == "nz" && (p.Count == 2 || p.Count == 3))
                        {
                            if (p[1].Length == 0) p[1].Append("''");
                            functionCallString = $"IIf(IsNull({p[0]}), {p[1]}, {p[0]})";
                        }
                        else
                        {
                            functionCallString = $"{functionName}({string.Join(", ", p).TrimEnd(',', ' ')})";
                        }

                        if (functionParams is null)
                        {
                            commands.Last().Text = commands.Last().Text + functionCallString + tokenSeparator;
                        }
                        else
                        {
                            functionParams.Last().Append(functionCallString + tokenSeparator);
                        }
                    }
                    else
                    {
                        enclosingCommandIndex = _CollectCommands(ctx, caretPosition, tokenSeparator, commandSeparatorTokenType, commands, enclosingCommandIndex, functionParams);
                        if (!(functionParams is null) && (ctx?.RuleIndex == MSAccessParser.RULE_function_name || ctx?.RuleIndex == MSAccessParser.RULE_param_expr))
                        {
                            functionParams.Last().Length -= tokenSeparator.Length;
                            functionParams.Add(new StringBuilder());
                        }
                    }
                }
            }
            return enclosingCommandIndex;
        }
    }
}
