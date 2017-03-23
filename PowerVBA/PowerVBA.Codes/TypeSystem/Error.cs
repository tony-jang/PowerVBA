using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerVBA.Codes.TypeSystem
{

    /// <summary>
    /// Enum that describes the type of an error.
    /// </summary>
    public enum ErrorType
    {
        Unknown,
        Error,
        Warning
    }

    /// <summary>
    /// Descibes an error during parsing.
    /// </summary>
    [Serializable]
    public class Error
    {
        readonly ErrorType errorType;
        readonly string message;
        readonly DomRegion region;

        /// <summary>
        /// The type of the error.
        /// </summary>
        public ErrorType ErrorType { get { return errorType; } }

        /// <summary>
        /// The error description.
        /// </summary>
        public string Message { get { return message; } }

        /// <summary>
        /// The region of the error.
        /// </summary>
        public DomRegion Region { get { return region; } }

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        /// <param name='region'>
        /// The region of the error.
        /// </param>
        public Error(ErrorType errorType, string message, DomRegion region)
        {
            this.errorType = errorType;
            this.message = message;
            this.region = region;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        /// <param name='location'>
        /// The location of the error.
        /// </param>
        public Error(ErrorType errorType, string message, TextLocation location)
        {
            this.errorType = errorType;
            this.message = message;
            this.region = new DomRegion(location, location);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        /// <param name='line'>
        /// The line of the error.
        /// </param>
        /// <param name='col'>
        /// The column of the error.
        /// </param>
        public Error(ErrorType errorType, string message, int line, int col) : this(errorType, message, new TextLocation(line, col))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ICSharpCode.NRefactory.TypeSystem.Error"/> class.
        /// </summary>
        /// <param name='errorType'>
        /// The error type.
        /// </param>
        /// <param name='message'>
        /// The description of the error.
        /// </param>
        public Error(ErrorType errorType, string message)
        {
            this.errorType = errorType;
            this.message = message;
            this.region = DomRegion.Empty;
        }
    }
}
