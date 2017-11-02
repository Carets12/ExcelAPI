/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.iesvdc.acceso.excelapi.excelapi;

/**
 *Esta clase contiene una excepción propia que hereda de Exception.
 * @author Daniel Sierra Ráez
 */
public class ExcelAPIException extends Exception {

    /**
     * Creates a new instance of <code>ExcelAPIException</code> without detail
     * message.
     */
    public ExcelAPIException() {
    }

    /**
     * Constructs an instance of <code>ExcelAPIException</code> with the
     * specified detail message.
     *
     * @param msg the detail message.
     */
    public ExcelAPIException(String msg) {
        super("ExcelAPIException:: " + msg);
    }
}
