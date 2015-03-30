/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Text;

    /// <summary>
    /// JSON text writer
    /// </summary>
    internal class JsonWriter : TextWriter
    {
        private const string Indentation = "  ";

        #region Member variables

        private Stream outStream;
        private bool shouldCloseStream;
        private bool prettyPrint;
        private bool writingStringValue = false;
        private byte[] smallBuffer = new byte[20];
        private char[] charBuffer = new char[4];
        private Queue<char> surrogateCharBuffer = new Queue<char>(4);

        private Stack<char> closures = new Stack<char>();
        private Stack<bool> closureHasMembers = new Stack<bool>();

        #endregion

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the <see cref="JsonWriter"/> class.
        /// </summary>
        /// <param name="outStream">The out stream.</param>
        /// <param name="prettyPrint">if set to <c>true</c> [pretty print].</param>
        public JsonWriter(Stream outStream, bool prettyPrint)
        {
            this.outStream = outStream;
            this.prettyPrint = prettyPrint;
            this.shouldCloseStream = false;
        }
        #endregion

        #region Dispose methods

        /// <summary>
        /// Releases the unmanaged resources used by the <see cref="T:System.IO.TextWriter"/> and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected override void Dispose(bool disposing)
        {
            if (this.shouldCloseStream)
            {
                this.outStream.Flush();
                this.outStream.Dispose();
            }

            base.Dispose(disposing);
        }
        #endregion

        #region Public methods
        /// <summary>
        /// Writes a character to the text stream.
        /// </summary>
        /// <param name="value">The character to write to the text stream.</param>
        /// <exception cref="T:System.ObjectDisposedException">
        /// The <see cref="T:System.IO.TextWriter"/> is closed.
        /// </exception>
        /// <exception cref="T:System.IO.IOException">
        /// An I/O error occurs.
        /// </exception>
        public override void Write(char value)
        {
            if (this.writingStringValue)
            {
                // Handle Surrogates and Supplementary Characters
                // Refer to: http://msdn.microsoft.com/en-us/library/windows/desktop/dd374069(v=vs.85).aspx
                if (char.IsHighSurrogate(value))
                {
                    this.surrogateCharBuffer.Enqueue(value);

                    if (this.surrogateCharBuffer.Count == 4)
                    {
                        Debug.Assert(false, "The number of surrogate characters for a single real character should be less than 4.");

                        // Still write characters out
                        this.WriteSurrogateChar();
                    }
                }
                else if (char.IsLowSurrogate(value))
                {
                    this.surrogateCharBuffer.Enqueue(value);
                    this.WriteSurrogateChar();
                }
                else
                {
                    // Non surrogate characters
                    if (this.surrogateCharBuffer.Count > 0)
                    {
                        Debug.Assert(false, "A low surrogate character is expected.");
                        this.WriteSurrogateChar();
                    }

                    if (value == '\r')
                    {
                        this.WriteInternal(@"\r");
                    }
                    else if (value == '\n')
                    {
                        this.WriteInternal(@"\n");
                    }
                    else if (value == '\t')
                    {
                        this.WriteInternal(@"\t");
                    }
                    else
                    {
                        this.WritePrintableChar(value);
                    }
                }
            }
            else
            {
                this.WritePrintableChar(value);
            }
        }

        /// <summary>
        /// Pushes object closure.
        /// </summary>
        public void PushObjectClosure()
        {
            this.AddingValue();
            this.closures.Push('}');
            this.closureHasMembers.Push(false);
            this.WriteInternal('{');
            this.WriteIndentation();
        }

        /// <summary>
        /// Pushes array closure.
        /// </summary>
        public void PushArrayClosure()
        {
            this.AddingValue();
            this.closures.Push(']');
            this.closureHasMembers.Push(false);
            this.WriteInternal('[');
            this.WriteIndentation();
        }

        /// <summary>
        /// Pops closure.
        /// </summary>
        public void PopClosure()
        {
            var popChar = this.closures.Pop();
            this.closureHasMembers.Pop();
            this.WriteIndentation();
            this.WriteInternal(popChar);
        }

        /// <summary>
        /// Writes quote.
        /// </summary>
        public void WriteQuote()
        {
            this.WriteInternal('"');
        }

        /// <summary>
        /// Writes key.
        /// </summary>
        /// <param name="key">The key.</param>
        public void WriteKey(string key)
        {
            if (this.closureHasMembers.Peek())
            {
                this.WriteInternal(',');
                this.WriteIndentation();
            }

            this.WriteQuote();
            this.Write(key);
            this.WriteQuote();

            if (this.prettyPrint)
            {
                this.Write(" : ");
            }
            else
            {
                this.WriteInternal(':');
            }
        }

        /// <summary>
        /// Writes value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(string value)
        {
            this.AddingValue();
            this.WriteQuote();
            this.writingStringValue = true;
            this.Write(value);
            this.writingStringValue = false;
            this.WriteQuote();
        }

        /// <summary>
        /// Writes bool value.
        /// </summary>
        /// <param name="value">if set to <c>true</c> [value].</param>
        public void WriteValue(bool value)
        {
            this.AddingValue();
            this.Write(value.ToString().ToLowerInvariant());
        }

        /// <summary>
        /// Writes long value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(long value)
        {
            this.AddingValue();
            this.Write(value);
        }

        /// <summary>
        /// Writes int value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(int value)
        {
            this.AddingValue();
            this.Write(value);
        }

        /// <summary>
        /// Writes an enum value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(Enum value)
        {
            this.WriteValue(EwsUtilities.SerializeEnum(value));
        }

        /// <summary>
        /// Writes DateTime value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(DateTime value)
        {
            this.WriteValue(EwsUtilities.DateTimeToXSDateTime(value));
        }

        /// <summary>
        /// Writes float value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(float value)
        {
            this.AddingValue();
            this.Write(value);
        }

        /// <summary>
        /// Writes double value.
        /// </summary>
        /// <param name="value">The value.</param>
        public void WriteValue(double value)
        {
            this.AddingValue();
            this.Write(value);
        }

        /// <summary>
        /// Writes null value.
        /// </summary>
        public void WriteNullValue()
        {
            this.AddingValue();
            this.Write("null");
        }
        #endregion

        #region Private methods

        /// <summary>
        /// Write printable char.
        /// </summary>
        /// <param name="value">The value.</param>
        private void WritePrintableChar(char value)
        {
            if (value == '"' || value == '\\')
            {
                this.WriteInternal('\\');
            }

            this.WriteInternal(value);
        }

        /// <summary>
        /// Internal writer.
        /// </summary>
        /// <param name="value">The value.</param>
        private void WriteInternal(char value)
        {
            this.charBuffer[0] = value;
            int bytesLength = this.Encoding.GetBytes(
                this.charBuffer,
                0,
                1,
                this.smallBuffer,
                0);

            this.outStream.Write(smallBuffer, 0, bytesLength);
        }

        /// <summary>
        /// Internal writer.
        /// </summary>
        /// <param name="value">The value.</param>
        private void WriteInternal(string value)
        {
            value.CopyTo(0, this.charBuffer, 0, value.Length);
            int bytesLength = this.Encoding.GetBytes(
                this.charBuffer,
                0,
                value.Length,
                this.smallBuffer,
                0);

            this.outStream.Write(smallBuffer, 0, bytesLength);
        }

        /// <summary>
        /// Surrogate character writer.
        /// </summary>
        private void WriteSurrogateChar()
        {
            this.surrogateCharBuffer.CopyTo(this.charBuffer, 0);
            int bytesLength = this.Encoding.GetBytes(
                this.charBuffer,
                0,
                this.surrogateCharBuffer.Count,
                this.smallBuffer,
                0);

            this.outStream.Write(smallBuffer, 0, bytesLength);
            this.surrogateCharBuffer.Clear();
        }

        /// <summary>
        /// Writes indentation.
        /// </summary>
        private void WriteIndentation()
        {
            if (this.prettyPrint)
            {
                this.WriteInternal('\r');
                this.WriteInternal('\n');

                for (int i = this.closures.Count - 1; i >= 0; i--)
                {
                    this.Write(JsonWriter.Indentation);
                }
            }
        }

        /// <summary>
        /// Adding a value.
        /// </summary>
        private void AddingValue()
        {
            if (this.closures.Count > 0)
            {
                if (this.closures.Peek() == ']' &&
                    this.closureHasMembers.Peek())
                {
                    this.WriteInternal(',');
                    this.WriteIndentation();
                }

                if (!this.closureHasMembers.Peek())
                {
                    this.closureHasMembers.Pop();
                    this.closureHasMembers.Push(true);
                }
            }
        }
        #endregion

        #region Public properties

        /// <summary>
        /// When overridden in a derived class, returns the <see cref="T:System.Text.Encoding"/> in which the output is written.
        /// </summary>
        /// <value></value>
        /// <returns>
        /// The Encoding in which the output is written.
        /// </returns>
        public override Encoding Encoding
        {
            get { return Encoding.UTF8; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether input stream should be closed when reader is closed.
        /// </summary>
        public bool ShouldCloseStream
        {
            get
            {
                return this.shouldCloseStream;
            }

            set
            {
                this.shouldCloseStream = value;
            }
        }
        #endregion
    }
}