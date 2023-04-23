import { styled } from 'config/stitches.config'

export const Footer = styled('div', {
  width: '100%',
  background: '$neutral1',
  marginTop: '48px',
  padding: '0px',
  span: {
    fontWeight: 600,
    display: 'block',
    marginBottom: '4px',
  },
})

export const Widgets = styled('div', {
  width: '100%',
  backgroundColor: '$gray100',
  margin: '0px',
  padding: '0 16px',
  borderTop: '1px solid #D7D7D7',
  borderBottom: '8px solid #840B09',
  ul: {
    padding: '0',
    li: {
      width: '25%',
      padding: '18px 0 50px 0',
      display: 'inline-block',
      justifyContent: '',
      alignItems: 'center',
      p: {
        color: '$gray700',
        fontSize: '$4',
        fontWeight: '$5',
      },
      '@bp3': {
        width: '100%',
      },
    },
  },
})

export const Copyright = styled('div', {
  width: '100%',
  height: '80px',
  background: '$neutral1',
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'center',
  fontSize: '$3',
  margin: '0px',
  padding: '0px',
  textAlign: 'center',
  a: {
    fontSize: '$3',
    color: '$primary1',
  },
})
