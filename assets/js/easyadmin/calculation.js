require('../../css/easyadmin/calculation.scss')

if (null !== document.querySelector('.easyadmin.edit-calculation')) {
    const calculator = document.querySelector('[name="calculation[calculator]"]')
    const update = () => {
        const calculatorClass = $(calculator).val()
        $('[data-calculator]').each((index, el) => {
            const enable = $(el).data('calculator') === calculatorClass
            $(el).closest('.form-group').toggle(enable)
            $(el).prop('disabled', !enable)
        })
    }

    $(calculator).on('change', update)
    update()
}
