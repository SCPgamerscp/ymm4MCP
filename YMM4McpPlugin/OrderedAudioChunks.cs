using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    // WASAPI callbacks supply reusable buffers. Copy each callback in arrival order.
    internal sealed class OrderedAudioChunks
    {
        internal const int DefaultMaxBytes = 64 * 1024 * 1024;
        private readonly List<byte[]> _chunks = new();
        private readonly object _gate = new();
        private readonly int _maxBytes;
        private int _size;
        private bool _truncated;

        internal OrderedAudioChunks(int maxBytes = DefaultMaxBytes)
        {
            if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            _maxBytes = maxBytes;
        }

        public bool IsTruncated
        {
            get { lock (_gate) return _truncated; }
        }

        public void Add(byte[] source, int length)
        {
            ArgumentNullException.ThrowIfNull(source);
            if (length < 0 || length > source.Length) throw new ArgumentOutOfRangeException(nameof(length));
            if (length == 0) return;
            lock (_gate)
            {
                if (_truncated) return;
                if (length > _maxBytes - _size)
                {
                    // Do not throw on the WASAPI callback thread or publish a
                    // truncated WAV as a successful capture.
                    _truncated = true;
                    return;
                }
                var chunk = new byte[length];
                Buffer.BlockCopy(source, 0, chunk, 0, length);
                _size += length;
                _chunks.Add(chunk);
            }
        }

        public byte[] ToArray()
        {
            lock (_gate)
            {
                var result = new byte[_size];
                int position = 0;
                foreach (var chunk in _chunks)
                {
                    Buffer.BlockCopy(chunk, 0, result, position, chunk.Length);
                    position += chunk.Length;
                }
                return result;
            }
        }
    }
}
