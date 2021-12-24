import { mergeStyleSets } from "@fluentui/react/lib/Styling";

export const classNames = mergeStyleSets({
    fullWidthControl: {
        width: '100%'
    },
    defaultButton: {
        width: '100%'
    },
    fileIcon: {
        fontSize: 20
    },
    wrapper: {
        height: '60vh',
        position: 'relative'
    },
    filter: {
        paddingBottom: 20,
        maxWidth: 300
    },
    header: {
        margin: 0
    },
    row: {
        display: 'inline-block'
    },
    commandbar: {
        padding: 0,
        root: {
            marginLeft: 0,
            marginTop: 0,
            padding: 0
        }
    },
    searchBox: {
        root: {
            width: '100%'
        }
    },
    stackItemStyles: {
        root: {
            alignItems: 'center',
            display: 'flex',
            height: 50,
            justifyContent: 'center',
        },
    },

});