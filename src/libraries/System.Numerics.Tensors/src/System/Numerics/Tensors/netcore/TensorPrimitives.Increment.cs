﻿// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Runtime.Intrinsics;

namespace System.Numerics.Tensors
{
    public static partial class TensorPrimitives
    {
        /// <summary>Computes the element-wise increment of numbers in the specified tensor.</summary>
        /// <param name="x">The tensor, represented as a span.</param>
        /// <param name="destination">The destination tensor, represented as a span.</param>
        /// <exception cref="ArgumentException">Destination is too short.</exception>
        /// <exception cref="ArgumentException"><paramref name="x"/> and <paramref name="destination"/> reference overlapping memory locations and do not begin at the same location.</exception>
        /// <remarks>
        /// <para>
        /// This method effectively computes <c><paramref name="destination" />[i] = <paramref name="x" />[i] + 1</c>.
        /// </para>
        /// </remarks>
        public static void Increment<T>(ReadOnlySpan<T> x, Span<T> destination)
            where T : IIncrementOperators<T>
        {
            if (typeof(T) == typeof(Half) && TryUnaryInvokeHalfAsInt16<T, IncrementOperator<float>>(x, destination))
            {
                return;
            }

            InvokeSpanIntoSpan<T, IncrementOperator<T>>(x, destination);
        }

        /// <summary>T.Increment(x)</summary>
        private readonly struct IncrementOperator<T> : IUnaryOperator<T, T>
            where T : IIncrementOperators<T>
        {
            public static bool Vectorizable => true;
            public static T Invoke(T x) => ++x;
            public static Vector128<T> Invoke(Vector128<T> x) => x + Vector128<T>.One;
            public static Vector256<T> Invoke(Vector256<T> x) => x + Vector256<T>.One;
            public static Vector512<T> Invoke(Vector512<T> x) => x + Vector512<T>.One;
        }
    }
}
