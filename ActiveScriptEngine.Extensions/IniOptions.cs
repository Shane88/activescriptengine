namespace ActiveXScriptLib.Extensions
{
   using System.Collections.Generic;

   public class IniOptions
   {
      internal string _iniFilePath;

      internal string _sectionName;

      internal readonly Dictionary<string, string> _tokens = new Dictionary<string, string>();

      public IniOptions IniFilePath(string iniFilePath)
      {
         this._iniFilePath = iniFilePath;
         return this;
      }

      public IniOptions Section(string sectionName)
      {
         this._sectionName = sectionName;
         return this;
      }

      public IniOptions AddPathToken(string tokenName, string tokenValue)
      {
         _tokens.Add(tokenName, tokenValue);
         return this;
      }
   }
}
