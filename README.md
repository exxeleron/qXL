## qXL


![Overview](../master/doc/img/qXL.png)

`qXL` is a dynamically loaded library using COM Interop interface to exchange data between Excel and kdb+ processes. Data serializing and de-serializing while communicating with kdb+ processes is realized with use of exxeleron’s [`qSharp`](https://github.com/exxeleron/qSharp) interface. 

The general functionality of the library covers:
- Request-reply mechanism for interaction with q processes
- Subscription interface for working with stream data (via RTD)
- Automatic type conversion between Excel and kdb+ types
- Support for Worksheet and VBA function calls via unified interface

## Getting started
### qXL download
Most recent `qXL` plugin for Excel can be downloaded from [here](../../releases).

### Documentation

Installation details, documentation and examples can be found [here](../master/doc/).
