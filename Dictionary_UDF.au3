#cs ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	Name..................: Dictionary_UDF
	Description...........: Dictionary basic functions: object that stores data key/item pairs
	Documentation.........: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object

    Author................: exorcistas@github.com
    Modified..............: 2019-12-30
	Version...............: v1.0
#ce ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#include-once

#Region FUNCTIONS_LIST
#cs	===================================================================================================================================
    _Dictionary_Create()
    _Dictionary_AddItem($_objDictionary, $_sKey, $_sValue)
    _Dictionary_SetItemValue($_objDictionary, $_sKey, $_sValue)
	_Dictionary_GetItemValue($_objDictionary, $_sKey)
	_Dictionary_KeyExists($_objDictionary, $_sKey)
	_Dictionary_ListItems($_objDictionary)
	_Dictionary_ListKeys($_objDictionary)
	_Dictionary_RemoveItem($_objDictionary, $_sKey)
	_Dictionary_RemoveAll($_objDictionary)
#ce	===================================================================================================================================
#EndRegion FUNCTIONS_LIST

#Region FUNCTIONS
	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_Create()
		Description .......: Creates Scripting.Dictionary object

		Return values .....: Success - Dictionary object
		                     Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-04
	#ce ===============================================================================================================================
    Func _Dictionary_Create()
        Local $_objDictionary = ObjCreate('Scripting.Dictionary')

        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)

        Return $_objDictionary
    EndFunc

	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_AddItem
		Description .......: Adds a key and item pair to a Dictionary object

		Return values .....: Success - TRUE
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-04
	#ce ===============================================================================================================================
    Func _Dictionary_AddItem($_objDictionary, $_sKey, $_sValue)
        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)
        If $_sKey = "" Then Return SetError(2, 0, False)

        If $_objDictionary.Exists($_sKey) Then Return SetError(3, 0, False)

        $_objDictionary.Add($_sKey, $_sValue)
        Return True
    EndFunc

	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_SetItemValue($_objDictionary, $_sKey, $_sValue)
		Description .......: Sets item value in Scripting.Dictionary object

		Return values .....: Success - TRUE
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-04
	#ce ===============================================================================================================================
    Func _Dictionary_SetItemValue($_objDictionary, $_sKey, $_sValue)
        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)
        If $_sKey = "" Then Return SetError(2, 0, False)

        If NOT $_objDictionary.Exists($_sKey) Then Return SetError(3, 0, False)

        $_objDictionary.Item($_sKey) = $_sValue
        Return True
    EndFunc

	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_GetItemValue($_objDictionary, $_sKey)
		Description .......: Gets item value from Scripting.Dictionary object

		Return values .....: Success - Item value
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-04
	#ce ===============================================================================================================================
    Func _Dictionary_GetItemValue($_objDictionary, $_sKey)
        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)
        If $_sKey = "" Then Return SetError(2, 0, False)

        If NOT $_objDictionary.Exists($_sKey) Then Return SetError(3, 0, False)

        Return $_objDictionary.Item($_sKey)
	EndFunc
	
	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_KeyExists($_objDictionary, $_sKey)
		Description .......: Returns True if a specified key exists in the Dictionary object; False if it does not.

		Return values .....: Success - TRUE or FALSE
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-30
	#ce ===============================================================================================================================
    Func _Dictionary_KeyExists($_objDictionary, $_sKey)
        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)
        If $_sKey = "" Then Return SetError(2, 0, False)

        Return $_objDictionary.Exists($_sKey)
	EndFunc
	
	#cs #FUNCTION# ====================================================================================================================
		Name...............:  _Dictionary_ListItems($_objDictionary)
		Description .......: Returns an array containing all the items in a Dictionary object

		Return values .....: Success - Array
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-30
	#ce ===============================================================================================================================
    Func _Dictionary_ListItems($_objDictionary)
        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)

        Return $_objDictionary.Items
	EndFunc
	
	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_ListKeys($_objDictionary)
		Description .......: Returns an array containing all existing keys in a Dictionary object

		Return values .....: Success - Array
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-30
	#ce ===============================================================================================================================
    Func _Dictionary_ListKeys($_objDictionary)
        If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)

        Return $_objDictionary.Keys
	EndFunc
	
	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_RemoveItem($_objDictionary, $_sKey)
		Description .......: Removes a key/item pair from a Dictionary object

		Return values .....: Success - TRUE
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-30
	#ce ===============================================================================================================================
    Func _Dictionary_RemoveItem($_objDictionary, $_sKey)
		If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)
		If $_sKey = "" Then Return SetError(2, 0, False)

		If NOT $_objDictionary.Exists($_sKey) Then Return SetError(3, 0, False)

		$_objDictionary.Remove($_sKey)
		Return True
	EndFunc

	#cs #FUNCTION# ====================================================================================================================
		Name...............: _Dictionary_RemoveAll($_objDictionary)
		Description .......: Removes all key, item pairs from a Dictionary object

		Return values .....: Success - TRUE
                             Failure - FALSE; @error

		Author ............: exorcistas@github.com
		Modified...........: 2019-12-30
	#ce ===============================================================================================================================
    Func _Dictionary_RemoveAll($_objDictionary)
		If NOT IsObj($_objDictionary) Then Return SetError(1, 0, False)

		$_objDictionary.RemoveAll
		Return True
	EndFunc
#EndRegion FUNCTIONS
