namespace ActiveXScriptLib.Extensions.Builder
{
   using System.Collections.Generic;

   public class IniOptions
   {
      public IniOptions()
      {
         ConfiguredTokens = new Dictionary<string, string>();
      }

      internal string ConfiguredIniFilePath { get; private set; }

      internal string ConfiguredSectionName { get; private set; }

      internal Dictionary<string, string> ConfiguredTokens { get; private set; }

      public IniOptions IniFilePath(string iniFilePath)
      {
         this.ConfiguredIniFilePath = iniFilePath;
         return this;
      }

      public IniOptions Section(string sectionName)
      {
         this.ConfiguredSectionName = sectionName;
         return this;
      }

      public ResolvesTo AddPathToken(string tokenName)
      {
         return new ResolvesTo(this, tokenName);
      }
   }

   public class ResolvesTo
   {
      private readonly IniOptions _options;
      private readonly string _tokenName;

      public ResolvesTo(IniOptions options, string tokenName)
      {
         this._options = options;
         this._tokenName = tokenName;
      }

      public IniOptions ThatResolvesTo(string tokenValue)
      {
         _options.ConfiguredTokens.Add(_tokenName, tokenValue);
         return _options;
      }
   }
}
