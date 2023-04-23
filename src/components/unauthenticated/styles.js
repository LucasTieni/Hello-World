import { styled } from 'config/stitches.config'

export const Unauthenticated = styled('div', {
  width: '100%',
  padding: 32,
  textAlign: 'center',
})

export const Button = styled('button', {
  borderRadius: 4,
  border: 'none',
  padding: '8px 16px',
  background: '$primary1',
  color: '#fff',
  fontSize: 16,
  lineHeight: 1,
  fontWeight: 500,
  cursor: 'pointer',
  marginTop: '16px',
})
