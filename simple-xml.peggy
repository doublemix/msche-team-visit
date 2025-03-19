{{
  let entityMap = {
    'amp': '&'
  }
}}

Main
	= Content*

Content
	= value:TextNode {
    	return { type: 'text', value }
    }
  / value:Tag {
      return { type: 'tag', ...value }
    }

TextNode
	= characters:Character+ {
    	return characters.join('')
    }

Character
	= [^<>&]
    / Entity

Entity
	= '&' value:$([^;]+) ';' {
    	let lookup = entityMap[value]
        if (!lookup) {
        	throw new Error('unknown entity: ' + value)
        }
    }

Tag
	= '<' tagName:Name _ attributes:Attributes _ 
    	rest:TagRest {
        
    	if (rest && tagName !== rest.endTagName) {
        	throw new Error('mismatched tag name: ' + rest.endTagName + ' does not match ' + tagName)
        }
        return { tagName, attributes, content: rest?.content ?? [] }
    }

TagRest
	= '/>' { return null }
    / '>' content:Content* '</' endTagName:Name '>' { return {content, endTagName} }

Attributes
	= Attribute|.., _|

Attribute
	= name:Name '=' value:AttributeValue {
    	return { name, value }
    }

AttributeValue
	= '"' characters:AttributeCharacter* '"' {
    	return characters.join('')
    }

AttributeCharacter
	= [^&"]
    / Entity

Name
	= @$([a-zA-Z:]+)

_
	= [ \r\t\n]*