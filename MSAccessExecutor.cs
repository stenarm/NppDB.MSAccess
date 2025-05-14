using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class MsAccessLexerErrorListener : ConsoleErrorListener<int>
    {
        public new static readonly MsAccessLexerErrorListener Instance = new MsAccessLexerErrorListener();

        public override void SyntaxError(TextWriter output, IRecognizer recognizer, 
            int offendingSymbol, int line, int col, string msg, RecognitionException e)
        {
            Console.WriteLine($@"LEXER ERROR: {e?.GetType().ToString() ?? ""}: {msg} ({line}:{col})");
        }
    }

    public class MsAccessParserErrorListener : BaseErrorListener
    {
        public IList<ParserError> Errors { get; } = new List<ParserError>();

        public override void SyntaxError(TextWriter output, IRecognizer recognizer, 
            IToken offendingSymbol, int line, int col, string msg, RecognitionException e)
        {
            Console.WriteLine($@"PARSER ERROR: {e?.GetType().ToString() ?? ""}: {msg} ({line}:{col})");
            Errors.Add(new ParserError
            {
                Text = msg,
                StartLine = line,
                StartColumn = col,
                StartOffset = offendingSymbol.StartIndex,
                StopOffset = offendingSymbol.StopIndex,
            }); 
        }
    }

    public sealed class MsAccessExecutor : ISqlExecutor
    {
        private Thread _execTh;
        private readonly Func<OleDbConnection> _connector;

        public MsAccessExecutor(Func<OleDbConnection> connector)
        {
            _connector = connector;
        }

        public ParserResult Parse(string sqlText, CaretPosition caretPosition)
        {
            var input = CharStreams.fromString(sqlText);

            var lexer = new MSAccessLexer(input);
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(MsAccessLexerErrorListener.Instance);

            CommonTokenStream tokens;
            try
            {
                tokens = new CommonTokenStream(lexer);
            }
            catch (Exception e)
            {
                Console.WriteLine($@"Lexer Exception: {e}");
                throw;
            }

            var parserErrorListener = new MsAccessParserErrorListener();
            var parser = new MSAccessParser(tokens);
            parser.RemoveErrorListeners();
            parser.AddErrorListener(parserErrorListener);
            try
            {
                var tree = parser.parse();
                var enclosingCommandIndex = tree.CollectCommands(caretPosition, " ", MSAccessParser.SCOL, out var commands);
                return new ParserResult
                {
                    Errors = parserErrorListener.Errors, 
                    Commands = commands.ToList<ParsedCommand>(), 
                    EnclosingCommandIndex = enclosingCommandIndex
                };
            }
            catch (Exception e)
            {
                Console.WriteLine($@"Parser Exception: {e}");
                throw;
            }
        }

        public SqlDialect Dialect => SqlDialect.MS_ACCESS;

        public void Execute(IList<string> sqlQueries, Action<IList<CommandResult>> callback)
        {
            _execTh = new Thread(new ThreadStart(
                delegate
                {
                    var results = new List<CommandResult>();
                    string lastSql = null;
                    try
                    {
                        using (var conn = _connector())
                        {
                            conn.Open();
                            foreach (var sql in sqlQueries)
                            {
                                if (string.IsNullOrWhiteSpace(sql)) continue;
                                lastSql = sql;

                                Console.WriteLine($@"SQL: <{sql}>");
                                var cmd = new OleDbCommand(sql, conn);
                                var rd = cmd.ExecuteReader();
                                var dt = new DataTable();
                                dt.Load(rd);
                                results.Add(new CommandResult {CommandText = sql, QueryResult = dt, RecordsAffected = rd.RecordsAffected});
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        results.Add(new CommandResult {CommandText = lastSql, Error = ex});
                        callback(results);
                        return;
                    }
                    callback(results);
                    _execTh = null;
                }));
            _execTh.IsBackground = true;
            _execTh.Start();
        }

        public bool CanExecute()
        {
            return !CanStop();
        }

        public void Stop()
        {
            if (!CanStop()) return;
            _execTh?.Abort();
            _execTh = null;
        }

        public bool CanStop()
        {
            // ReSharper disable once NonConstantEqualityExpressionHasConstantResult
            return _execTh != null && (_execTh.ThreadState & ThreadState.Running) != 0;
        }

    }
}
