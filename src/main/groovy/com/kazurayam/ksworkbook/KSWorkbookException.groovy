package com.kazurayam.ksworkbook

class KSWorkbookException extends Exception {
    
    //
    private static final long serialVersionUID = 1L;

    KSWorkbookException(String msg){
        super(msg)
    }

    KSWorkbookException(String msg, Throwable cause){
        super(msg, cause)
    }

    KSWorkbookException(Throwable cause) {
        super(cause)
    }

    KSWorkbookException() {
        super()
    }
}
