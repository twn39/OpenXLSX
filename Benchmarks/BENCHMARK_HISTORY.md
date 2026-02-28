# OpenXLSX Benchmark History

This file tracks the performance of OpenXLSX across major refactors and feature updates.

## 2026-02-28: Post-Security & Architecture Overhaul (Baseline)

**System Environment:**
- **OS**: Darwin (macOS)
- **Architecture**: aarch64 (Apple Silicon)
- **Compiler**: AppleClang 17.0.0.17000603
- **Build Type**: Release
- **LTO**: Enabled (where supported)

**Key Components:**
- **Microsoft GSL**: v4.2.1
- **minizip-ng**: v4.0.8
- **zlib-ng**: v2.2.4 (SIMD Optimized)
- **pugixml**: v1.14 (Internal)
- **Catch2**: v3.13.0

### Results (Mean Execution Time)

| Operation | Performance (Mean) | Std Dev | Notes |
| :--- | :--- | :--- | :--- |
| **Write Strings** | 665.909 ms | 5.27 ms | In-memory XML generation |
| **Write Integers** | 721.972 ms | 4.67 ms | In-memory XML generation |
| **Write Floats** | 740.798 ms | 3.33 ms | In-memory XML generation |
| **Write Bools** | 641.021 ms | 8.86 ms | In-memory XML generation |
| **Read Strings** | 150.742 ms | 0.62 ms | XML parsing + Cache lookup |
| **Read Integers** | 114.953 ms | 2.34 ms | XML parsing + numeric conversion |

### Raw Catch2 Output Summary
```text
benchmark name                       samples       iterations    mean          std dev
---------------------------------------------------------------------------------------
Write Strings                                  100             1     665.909 ms    5.27569 ms
Write Integers                                 100             1     721.972 ms    4.67453 ms
Write Floats                                   100             1     740.798 ms    3.33865 ms
Write Bools                                    100             1     641.021 ms    8.86979 ms
Read Strings                                   100             1     150.742 ms    626.27 us
Read Integers                                  100             1     114.953 ms    2.34031 ms
```

### Observations
- **Baseline established**: This is the first official benchmark after the full integration of GSL and the migration to minizip-ng.
- **Safety Overheads**: The impact of GSL's `span` and `narrow` is minimal, with reading operations remaining well under 200ms.
- **Next Steps**: Future optimizations focusing on floating-point to string conversion could reduce the `Write Floats` time.
