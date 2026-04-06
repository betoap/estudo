readonly sections = signal<OfferSection[]>([
    {
      type: 'negotiated',
      title: 'Negociadas',
      description:
        'Preços diferenciados aprovados no sistema Visão Cliente. Não podem ter alteração e seguem para a formalização do cliente, sem aprovações da mesa de crédito.',
      offers: [
        {
          id: 1,
          badge: 'Taxa negociada',
          title: 'Aplicação Financeira',
          subtitle: 'Percentual da garantia: 100%',
          icon: '≋',
          availabilityLabel: 'Disponibilidade',
          highlightValue: 'R$ 120.000,00',
          contractLabel: 'Prazo do contrato',
          contractValue: '12x de R$ 11.240,00',
          secondaryInfo: [
            {
              label: 'Seguro',
              value: 'Incluído R$ 25,47 a.m. do seguro',
              icon: '🛡️',
              muted: true,
            },
          ],
          rightInfo: [
            { label: 'Data da 1ª parcela', value: '14 jan 2026' },
            { label: 'Spread ao ano', value: '2,67% a.m.' },
            { label: 'Taxa de juros', value: '1,99% a.m. / 26,67% a.a.' },
          ],
          actionLabel: 'Contratar',
          actionVariant: 'primary',
        },
      ],
    },
    {
      type: 'preApproved',
      title: 'Pré-aprovadas',
      description:
        'Escolha a garantia e comece uma simulação. Ofertas pré-aprovadas podem passar por mais etapas de aprovação e isso pode impactar o tempo de contratação.',
      offers: [
        {
          id: 2,
          badge: 'Pré-aprovado',
          title: 'Devedor solidário',
          subtitle: 'Percentual da garantia: 100%',
          icon: '👥',
          availabilityLabel: 'Disponibilidade',
          highlightValue: 'R$ 85.000,00',
          contractLabel: 'Prazo do contrato',
          contractValue: 'Entre 6 e 60 meses',
          rightInfo: [
            {
              label: 'Transferência de limite',
              value: 'Entre R$ 5.000,00 e R$ 85.000,00',
            },
          ],
          actionLabel: 'Simular',
          actionVariant: 'outline',
        },
        {
          id: 3,
          badge: 'Pré-aprovado',
          title: 'Aplicação Financeira',
          subtitle: 'Percentual da garantia: 100%',
          icon: '≋',
          availabilityLabel: 'Disponibilidade',
          highlightValue: 'R$ 230.000,00',
          contractLabel: 'Prazo do contrato',
          contractValue: 'Entre 6 e 60 meses',
          rightInfo: [
            {
              label: 'Investimentos',
              value: 'Dos sócios e da empresa, exceto conta conjunta',
            },
          ],
          actionLabel: 'Simular',
          actionVariant: 'outline',
        },
      ],
    },
  ]);
}
