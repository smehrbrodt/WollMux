package de.muenchen.allg.itd51.wollmux.former;

import java.util.regex.Pattern;

import javax.swing.text.Segment;

import org.fife.ui.rsyntaxtextarea.AbstractTokenMaker;
import org.fife.ui.rsyntaxtextarea.RSyntaxUtilities;
import org.fife.ui.rsyntaxtextarea.Token;
import org.fife.ui.rsyntaxtextarea.TokenMap;

public class ConfigTokenMaker extends AbstractTokenMaker
{
  private final static Pattern IDENT = Pattern.compile("[a-zA-Z][_a-zA-Z0-9]*");

  @Override
  public Token getTokenList(Segment text, int initialTokenType, int startOffset)
  {
    resetTokenList();

    int offset = text.offset;
    int count = text.count;
    int end = offset + count;
    int newStartOffset = startOffset - offset;

    int tokenStart = offset;
    int tokenType = initialTokenType;

    for (int i = offset; i < end; i++)
    {
      char c = text.array[i];

      switch (tokenType)
      {
        case Token.NULL:
          tokenStart = i;

          switch (c)
          {
            case ' ':
            case '\t':
              tokenType = Token.WHITESPACE;
              break;
            case '"':
              tokenType = Token.ERROR_STRING_DOUBLE;
              break;
            case '\'':
              tokenType = Token.ERROR_CHAR;
              break;
            case '(':
            case ')':
            case ',':
            case ';':
              tokenType = Token.SEPARATOR;
              break;
            case '#':
              tokenType = Token.COMMENT_EOL;
              break;
            default:
              if (RSyntaxUtilities.isLetter(c) || c == '%')
              {
                tokenType = Token.IDENTIFIER;
              }
              else
              {
                tokenType = Token.ERROR_IDENTIFIER;
              }
          }
          break;
        case Token.WHITESPACE:
          switch (c)
          {
            case ' ':
            case '\t':
              break;
            case '"':
              addToken(text, tokenStart, i - 1, Token.WHITESPACE, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.ERROR_STRING_DOUBLE;
              break;
            case '\'':
              addToken(text, tokenStart, i - 1, Token.WHITESPACE, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.ERROR_CHAR;
              break;
            case '(':
            case ')':
            case ',':
            case ';':
              addToken(text, tokenStart, i - 1, Token.WHITESPACE, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.SEPARATOR;
              break;
            case '#':
              tokenType = Token.COMMENT_EOL;
              break;
            default:
              addToken(text, tokenStart, i - 1, Token.WHITESPACE, newStartOffset
                + tokenStart);
              tokenStart = i;
              if (RSyntaxUtilities.isLetter(c) || c == '%')
              {
                tokenType = Token.IDENTIFIER;
              }
              else
              {
                tokenType = Token.ERROR_IDENTIFIER;
              }
          }
          break;
        case Token.IDENTIFIER:
          switch (c)
          {

            case ' ':
            case '\t':
              addToken(text, tokenStart, i - 1, Token.IDENTIFIER, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.WHITESPACE;
              break;
            case '(':
            case ')':
            case ',':
            case ';':
              addToken(text, tokenStart, i - 1, Token.IDENTIFIER, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.SEPARATOR;
              break;
            case '"':
              addToken(text, tokenStart, i - 1, Token.IDENTIFIER, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.ERROR_STRING_DOUBLE;
              break;
            case '\'':
              addToken(text, tokenStart, i - 1, Token.IDENTIFIER, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.ERROR_CHAR;
              break;
            default:
              if (RSyntaxUtilities.isLetterOrDigit(c) || c == '_')
              {
                break;
              } 
              else
              {
                tokenType = Token.ERROR_IDENTIFIER;
              }
          }
          break;
        case Token.ERROR_STRING_DOUBLE:
          if (c == '"')
          {
            addToken(text, tokenStart, i, Token.LITERAL_STRING_DOUBLE_QUOTE,
              newStartOffset + tokenStart);
            tokenStart = i + 1;
            tokenType = Token.NULL;
          }
          break;
        case Token.ERROR_CHAR:
          if (c == '\'')
          {
            addToken(text, tokenStart, i, Token.LITERAL_CHAR,
              newStartOffset + tokenStart);
            tokenStart = i + 1;
            tokenType = Token.NULL;
          }
          break;
        case Token.ERROR_IDENTIFIER:
          if (c == ' ' || c == '\t')
          {
            addToken(text, tokenStart, i - 1, Token.ERROR_IDENTIFIER, newStartOffset
              + tokenStart);
            tokenStart = i;
            tokenType = Token.WHITESPACE;
          }
          break;
        case Token.SEPARATOR:
          switch (c)
          {
            case ' ':
            case '\t':
              break;
            case '"':
              addToken(text, tokenStart, i - 1, Token.SEPARATOR, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.ERROR_STRING_DOUBLE;
              break;
            case '\'':
              addToken(text, tokenStart, i - 1, Token.SEPARATOR, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.ERROR_CHAR;
              break;
            case '(':
            case ')':
            case ',':
            case ';':
              addToken(text, tokenStart, i - 1, Token.SEPARATOR, newStartOffset
                + tokenStart);
              tokenStart = i;
              tokenType = Token.SEPARATOR;
              break;
            default:
              addToken(text, tokenStart, i - 1, Token.SEPARATOR, newStartOffset
                + tokenStart);
              tokenStart = i;
              if (RSyntaxUtilities.isLetter(c))
              {
                tokenType = Token.IDENTIFIER;
              }
              else
              {
                tokenType = Token.ERROR_IDENTIFIER;
              }
          }
          break;
        case Token.COMMENT_EOL:
          break;
        default:
          addToken(text, tokenStart, i, tokenType, newStartOffset
            + tokenStart);
          tokenStart = i + 1;
          tokenType = Token.NULL;
      }
    }

    switch (tokenType)
    {
      case Token.NULL:
        addNullToken();
        break;
      default:
        addToken(text, tokenStart, end - 1, tokenType, newStartOffset + tokenStart);
        addNullToken();

    }
    return firstToken;
  }

  @Override
  public TokenMap getWordsToHighlight()
  {
    TokenMap tokenMap = new TokenMap();

    tokenMap.put("%include", Token.RESERVED_WORD);

    return tokenMap;
  }

  @Override
  public void addToken(Segment segment, int start, int end, int tokenType,
      int startOffset)
  {
    if (tokenType == Token.IDENTIFIER)
    {
      int value = wordsToHighlight.get(segment, start, end);
      if (value != -1)
      {
        tokenType = value;
      }
    }
    super.addToken(segment, start, end, tokenType, startOffset);
  }
}
