using System;
using System.Collections.Generic;

namespace YMM4McpPlugin
{
    // WASAPI callbacks supply reusable buffers. Copy each callback in arrival order.
    internal sealed class OrderedAudioChunks
    {
        private readonly List<byte[]> _chunks = new();
        private readonly object _gate = new();
        private int _size;

        public void Add(byte[] source, int length)
        {
            if (length < 0 || length > source.Length) throw new ArgumentOutOfRangeException(nameof(length));
            var chunk = new byte[length];
            Buffer.BlockCopy(source, 0, chunk, 0, length);
            lock (_gate)
            {
                _size = checked(_size + length);
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
