<!--#include file="./phpvbs.asp"-->
<%

'/**
' * Convenience method for htmlspecialchars.
' *
' * @param string text Text to wrap through htmlspecialchars
' * @return string Wrapped text
' */
    Function h(text)
        If isArray(text) Then
             h = array_map("h",text)
        ElseIf isObject(text) Then
            set h = array_map("h",text)
        Else
            h = htmlspecialchars(text)
        End If
    End Function

'/**
' * Constructs associative array from pairs of arguments.
' *
' * Example:
' * <code>
' * aa(array('a','b'))
' * </code>
' *
' * Would return:
' * <code>
' * array('a'=>'b')
' * </code>
' *
' * @return array Associative array
' */
    Function aa(args)

        Dim argc,i,a
        set a = Server.CreateObject("Scripting.Dictionary")
        argc = uBound(args)
        for i = 0 to argc

            If i + 1 <= argc Then
                a.add args(i),args(i + 1)
            Else
                a.add args(i), null
            End If
            i = i + 1
        Next

        set aa = a
    End Function

'/**
' * Convenience method for echo().
' *
' * @param string text String to echo
' */
    Function e(text)
        echo(text)
    End Function

'/**
' * Convenience method for strtolower().
' *
' * @param string str String to lowercase
' * @return string Lowercased string
' */
    Function low(str)
        low = strtolower(str)
    End Function

'/**
' * Convenience method for strtoupper().
' *
' * @param string str String to uppercase
' * @return string Uppercased string
' */
    Function up(str)
        up = strtoupper(str)
    End Function

'/**
' * Convenience method for str_replace().
' *
' * @param string search String to be replaced
' * @param string replace String to insert
' * @param string subject String to search
' * @return string Replaced string
' */
    Function r(search, replace, subject)
        r = str_replace(search, replace, subject)
    End Function

'/**
' * Print_r convenience function, which prints out <PRE> tags around
' * the output of given array. Similar to debug().
' *
' * @see	debug()
' * @param array var Variable to print out
' */
    Function pr(var)
        Response.Write("<pre>")
        print_r var,0
        Response.Write("</pre>")
    End Function

'/**
' * is_email
' *
' * @param string user_email mailaddress String
' */

    Function is_email(user_email)

        is_email = false
        If len(user_email) <= 0 Then Exit Function

        Dim chars
        chars = "/^([a-z0-9+_]|¥-|¥.)+@(([a-z0-9_]|¥-)+¥.)+[a-z]{2,6}$/i"

        If inStr(user_email,"@") > 0 and inStr(user_email,".") > 0 Then
            If preg_match(chars,user_email,"","","") Then
                is_email = true
            End If
        End If

    End Function

'/**
' * sprintfの拡張
' * ASPでは引数を動的に変更できないため、配列で指定
' *
' * @param array myAry1
' * @param array myAry2
' */

    Function array_sprintf(myAry1,myAry2)

        Dim value

        If Not isArray(myAry1) OR Not isArray(myAry2) Then
            value = sprintf(myAry1,myAry2)
        Else
            Dim i

            For i = 0 to uBound(myAry1)
                value = value & sprintf(myAry1(i),myAry2(i))
            Next

        End If

        array_sprintf = value

    End Function

%>
