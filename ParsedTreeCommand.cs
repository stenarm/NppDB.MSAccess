using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NppDB.Comm;

namespace NppDB.MSAccess
{
    public class ParsedTreeCommand : ParsedCommand
    {
        public IParseTree Context { get; set; }
        
        public void AddWarning(ParserRuleContext ctx, ParserMessageType type)
        {
            var startToken = ctx.Start;
            var stopToken = ctx.Start;
            var warning = new ParserWarning
            {
                Type = type,
                Text = nameof(type),
                StartLine = startToken.Line,
                StartColumn = startToken.Column,
                StartOffset = startToken.StartIndex,
                StopLine = stopToken.Line,
                StopColumn = stopToken.Column,
                StopOffset = stopToken.StopIndex,
            };
            Warnings.Add(warning);
        }
        
        public void AddWarning(IToken token, ParserMessageType type)
        {
            var warning = new ParserWarning
            {
                Type = type,
                Text = nameof(type),
                StartLine = token.Line,
                StartColumn = token.Column,
                StartOffset = token.StartIndex,
                StopLine = token.Line,
                StopColumn = token.Column,
                StopOffset = token.StopIndex == token.StartIndex ? token.StopIndex + 1 : token.StopIndex,
            };
            Warnings.Add(warning);
        }
    }
}
